VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "COMCTL32.OCX"
Begin VB.Form Formsetting 
   Appearance      =   0  '����
   BorderStyle     =   1  '��u�T�w
   Caption         =   "UnlightVBE-�i���]�w"
   ClientHeight    =   8730
   ClientLeft      =   6330
   ClientTop       =   2535
   ClientWidth     =   17670
   BeginProperty Font 
      Name            =   "�L�n������"
      Size            =   12
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Formsetting.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   17670
   Begin VB.Frame ���_�t�� 
      Caption         =   "�t��"
      Height          =   5895
      Left            =   9480
      TabIndex        =   135
      Top             =   2160
      Width           =   9015
      Begin VB.ComboBox cbsimilarlevel 
         BeginProperty Font 
            Name            =   "�L�n������"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4560
         TabIndex        =   150
         Text            =   "Combo1"
         Top             =   2880
         Width           =   975
      End
      Begin VB.CheckBox chkusesimilarlevel 
         Caption         =   "�H�������H���ɥH�ۦ������Ŷi���ԡG"
         BeginProperty Font 
            Name            =   "�L�n������"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   149
         Top             =   2880
         Width           =   6495
      End
      Begin VB.TextBox �ۭqAI��P�i�� 
         Appearance      =   0  '����
         BorderStyle     =   0  '�S���ؽu
         Enabled         =   0   'False
         Height          =   300
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   146
         Text            =   "7"
         Top             =   5040
         Width           =   375
      End
      Begin VB.CheckBox chksetcomaipagenum 
         Caption         =   "�ۭqAI�p�⤧��P�i�ơG        �i"
         BeginProperty Font 
            Name            =   "�L�n������"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   145
         Top             =   5040
         Width           =   6495
      End
      Begin VB.CheckBox chkusenewinterface 
         Caption         =   "�ϥηs�����q��ܤ���"
         BeginProperty Font 
            Name            =   "�L�n������"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   144
         Top             =   1440
         Width           =   2895
      End
      Begin VB.CheckBox chkusenewaipersonauto 
         Caption         =   "�ϥΪ̤�ϥΡu����P�_���H�u���zAI�v�i����"
         BeginProperty Font 
            Name            =   "�L�n������"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   141
         Top             =   2400
         Width           =   6495
      End
      Begin VB.CheckBox checktestpersondown 
         Caption         =   "���ռҦ�(�}�Ҧۭq�v�l�y��)"
         BeginProperty Font 
            Name            =   "�L�n������"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   138
         Top             =   1920
         Width           =   6135
      End
      Begin VB.CheckBox chkusenewpage 
         Caption         =   "�ϥηs���H�����ܤƤ���ʥd�P��"
         BeginProperty Font 
            Name            =   "�L�n������"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   137
         Top             =   960
         Width           =   6495
      End
      Begin VB.CheckBox chkusenewai 
         Caption         =   "�q����ϥΡu����P�_���H�u���zAI�v�i����"
         BeginProperty Font 
            Name            =   "�L�n������"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   136
         Top             =   480
         Width           =   6495
      End
      Begin VB.Label Label8 
         Caption         =   "�H����ܥ��t2���Ŭ��H�����W�U���A�Y3���H�����ⳣ���ŦX�d��̫h�H��3�����G+�H�����Ŭ��D�C"
         BeginProperty Font 
            Name            =   "�L�n������"
            Size            =   8.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   151
         Top             =   3240
         Width           =   8175
      End
      Begin VB.Label Label7 
         Caption         =   "�Ъ`�N�G�Y�N�p��i�Ƴ]�w�L���i��|�ɭP�{���S���^���C"
         BeginProperty Font 
            Name            =   "�L�n������"
            Size            =   8.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   147
         Top             =   5400
         Width           =   5655
      End
   End
   Begin VB.Frame �ƥ�d_�q�� 
      Caption         =   "�ƥ�d�s��(�q����)"
      Height          =   5895
      Left            =   9480
      TabIndex        =   51
      Top             =   2040
      Visible         =   0   'False
      Width           =   9015
      Begin VB.CheckBox persontgrecom 
         Caption         =   "��uUnlight�ƥ�d�W�h"
         BeginProperty Font 
            Name            =   "�L�n������"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   108
         Top             =   360
         Value           =   1  '�֨�
         Width           =   2535
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  '����
         BorderStyle     =   0  '�S���ؽu
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5160
         ScaleHeight     =   615
         ScaleWidth      =   3615
         TabIndex        =   103
         Top             =   240
         Width           =   3615
         Begin VB.OptionButton persontgruoncom 
            Caption         =   "�L"
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   107
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton persontgruoncom 
            Caption         =   "�ۭq"
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   120
            TabIndex        =   106
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton persontgruoncom 
            Caption         =   "��̤ܳj��"
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   1080
            TabIndex        =   105
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton persontgruoncom 
            Caption         =   "�H��"
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   4
            Left            =   1080
            TabIndex        =   104
            Top             =   220
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1695
         Left            =   120
         TabIndex        =   82
         Top             =   720
         Width           =   8535
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   2400
            TabIndex        =   88
            Top             =   240
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   2400
            TabIndex        =   87
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   3
            Left            =   2400
            TabIndex        =   86
            Top             =   1200
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   4
            Left            =   6000
            TabIndex        =   85
            Top             =   240
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   6
            Left            =   6000
            TabIndex        =   83
            Top             =   1200
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   5
            Left            =   6000
            TabIndex        =   84
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label persontgcom 
            Caption         =   "1:"
            Height          =   375
            Index           =   1
            Left            =   2040
            TabIndex        =   96
            Top             =   240
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "2:"
            Height          =   375
            Index           =   2
            Left            =   2040
            TabIndex        =   95
            Top             =   720
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "3:"
            Height          =   375
            Index           =   3
            Left            =   2040
            TabIndex        =   94
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "4:"
            Height          =   375
            Index           =   4
            Left            =   5640
            TabIndex        =   93
            Top             =   240
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "5:"
            Height          =   375
            Index           =   5
            Left            =   5640
            TabIndex        =   92
            Top             =   720
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "6:"
            Height          =   375
            Index           =   6
            Left            =   5640
            TabIndex        =   91
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label personlevelcom 
            BackStyle       =   0  '�z��
            Caption         =   "LV"
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
            TabIndex        =   90
            Top             =   720
            Width           =   495
         End
         Begin VB.Label personnamecom 
            Caption         =   "XXX"
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   600
            TabIndex        =   89
            Top             =   720
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1695
         Left            =   120
         TabIndex        =   67
         Top             =   2400
         Width           =   8535
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   12
            Left            =   6000
            TabIndex        =   73
            Top             =   1200
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   11
            Left            =   6000
            TabIndex        =   72
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   10
            Left            =   6000
            TabIndex        =   71
            Top             =   240
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   9
            Left            =   2400
            TabIndex        =   70
            Top             =   1200
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   8
            Left            =   2400
            TabIndex        =   69
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   7
            Left            =   2400
            TabIndex        =   68
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label personnamecom 
            Caption         =   "XXX"
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   600
            TabIndex        =   81
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label personlevelcom 
            BackStyle       =   0  '�z��
            Caption         =   "LV"
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
            Left            =   120
            TabIndex        =   80
            Top             =   720
            Width           =   495
         End
         Begin VB.Label persontgcom 
            Caption         =   "6:"
            Height          =   375
            Index           =   12
            Left            =   5640
            TabIndex        =   79
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "5:"
            Height          =   375
            Index           =   11
            Left            =   5640
            TabIndex        =   78
            Top             =   720
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "4:"
            Height          =   375
            Index           =   10
            Left            =   5640
            TabIndex        =   77
            Top             =   240
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "3:"
            Height          =   375
            Index           =   9
            Left            =   2040
            TabIndex        =   76
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "2:"
            Height          =   375
            Index           =   8
            Left            =   2040
            TabIndex        =   75
            Top             =   720
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "1:"
            Height          =   375
            Index           =   7
            Left            =   2040
            TabIndex        =   74
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1695
         Left            =   120
         TabIndex        =   52
         Top             =   4080
         Width           =   8535
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   18
            Left            =   6000
            TabIndex        =   58
            Top             =   1200
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   17
            Left            =   6000
            TabIndex        =   57
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   16
            Left            =   6000
            TabIndex        =   56
            Top             =   240
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   15
            Left            =   2400
            TabIndex        =   55
            Top             =   1200
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   14
            Left            =   2400
            TabIndex        =   54
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox personcom 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   13
            Left            =   2400
            TabIndex        =   53
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label personnamecom 
            Caption         =   "XXX"
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   600
            TabIndex        =   66
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label personlevelcom 
            BackStyle       =   0  '�z��
            Caption         =   "LV"
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
            Left            =   120
            TabIndex        =   65
            Top             =   720
            Width           =   495
         End
         Begin VB.Label persontgcom 
            Caption         =   "6:"
            Height          =   375
            Index           =   18
            Left            =   5640
            TabIndex        =   64
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "5:"
            Height          =   375
            Index           =   17
            Left            =   5640
            TabIndex        =   63
            Top             =   720
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "4:"
            Height          =   375
            Index           =   16
            Left            =   5640
            TabIndex        =   62
            Top             =   240
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "3:"
            Height          =   375
            Index           =   15
            Left            =   2040
            TabIndex        =   61
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "2:"
            Height          =   375
            Index           =   14
            Left            =   2040
            TabIndex        =   60
            Top             =   720
            Width           =   255
         End
         Begin VB.Label persontgcom 
            Caption         =   "1:"
            Height          =   375
            Index           =   13
            Left            =   2040
            TabIndex        =   59
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Label personwagcom 
         Caption         =   "(���⦳�H���ɵL�k�ۭq)"
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
         Left            =   3000
         TabIndex        =   109
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame �@��]�w_��� 
      Caption         =   "�@��]�w"
      Height          =   6015
      Left            =   120
      TabIndex        =   111
      Top             =   2040
      Width           =   9015
      Begin VB.Frame ��L�]�w 
         Caption         =   "��L�]�w"
         Height          =   1215
         Left            =   120
         TabIndex        =   129
         Top             =   4560
         Width           =   8775
         Begin VB.TextBox �j�ð��Ҧ��ﶵ_�P�� 
            Appearance      =   0  '����
            BorderStyle     =   0  '�S���ؽu
            Enabled         =   0   'False
            Height          =   375
            Left            =   3240
            MaxLength       =   2
            TabIndex        =   148
            Text            =   "17"
            Top             =   720
            Width           =   375
         End
         Begin VB.TextBox �D�ԼҦ��ﶵ_�P�� 
            Appearance      =   0  '����
            BorderStyle     =   0  '�S���ؽu
            Enabled         =   0   'False
            Height          =   375
            Left            =   3240
            MaxLength       =   2
            TabIndex        =   132
            Text            =   "4"
            Top             =   360
            Width           =   375
         End
         Begin VB.CheckBox �D�ԼҦ��ﶵ 
            Caption         =   "�D�ԼҦ��]��Թ��h�o        �i�P�^(Max:30)"
            Height          =   300
            Left            =   240
            TabIndex        =   133
            Top             =   360
            Width           =   5655
         End
         Begin VB.CheckBox checktest 
            Caption         =   "���ռҦ�(Admin)"
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
            Left            =   5880
            TabIndex        =   131
            Top             =   360
            Width           =   2535
         End
         Begin VB.CheckBox �j�ð��ﶵ 
            Caption         =   "�j�ð��Ҧ��]���訤��o        �i�P�AHP=99�^ (Min:1)"
            Height          =   375
            Left            =   240
            TabIndex        =   130
            Top             =   720
            Width           =   6255
         End
      End
      Begin VB.Frame ��L 
         Caption         =   "��L"
         Height          =   1215
         Left            =   120
         TabIndex        =   125
         Top             =   3240
         Width           =   8775
         Begin VB.TextBox ckendturnnum 
            Height          =   420
            Left            =   1080
            TabIndex        =   126
            Text            =   "18"
            Top             =   600
            Width           =   495
         End
         Begin VB.CheckBox chkpersonvsmode 
            Caption         =   "���ԼҦ�"
            Height          =   300
            Left            =   240
            TabIndex        =   128
            Top             =   280
            Width           =   1575
         End
         Begin VB.CheckBox ckendturn 
            Caption         =   "���           �^�X"
            Height          =   375
            Left            =   240
            TabIndex        =   127
            Top             =   600
            Width           =   1935
         End
      End
      Begin VB.Frame �I���Ϥ� 
         Caption         =   "�I���Ϥ��έ���"
         Height          =   2895
         Left            =   120
         TabIndex        =   112
         Top             =   360
         Width           =   8775
         Begin VB.CommandButton Command2 
            Caption         =   "�q�ɮ�..."
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
            Left            =   4080
            TabIndex        =   139
            Top             =   1920
            Width           =   1095
         End
         Begin VB.ComboBox BGM��� 
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
            Left            =   6120
            TabIndex        =   117
            Text            =   "Combo1"
            Top             =   2280
            Width           =   2535
         End
         Begin VB.ComboBox ��Ԧa�Ͽ�� 
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   435
            Left            =   2880
            TabIndex        =   116
            Text            =   "Combo2"
            Top             =   2280
            Width           =   2295
         End
         Begin VB.CommandButton Command1 
            Caption         =   "�q�ɮ�..."
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
            Left            =   7560
            TabIndex        =   115
            Top             =   1920
            Width           =   1095
         End
         Begin VB.CheckBox ckbgmmute 
            Caption         =   "�R��"
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7920
            TabIndex        =   114
            Top             =   480
            Width           =   735
         End
         Begin VB.CheckBox cksemute 
            Caption         =   "�R��"
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7920
            TabIndex        =   113
            Top             =   960
            Width           =   735
         End
         Begin MSComDlg.CommonDialog cdgBGMchooce 
            Left            =   7080
            Top             =   1800
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            CancelError     =   -1  'True
            DialogTitle     =   "UnlightVBE-BGM���-�}���ɮ�"
         End
         Begin MSComDlg.CommonDialog cdgMAPchooce 
            Left            =   3600
            Top             =   1800
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            CancelError     =   -1  'True
            DialogTitle     =   "UnlightVBE-��Ԧa�Ͽ��-�}���ɮ�"
         End
         Begin ComctlLib.Slider sdrbgm 
            Height          =   495
            Left            =   4680
            TabIndex        =   142
            Top             =   360
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   873
            _Version        =   327682
            LargeChange     =   10
            Min             =   1
            Max             =   100
            SelStart        =   50
            Value           =   50
         End
         Begin ComctlLib.Slider sdrse 
            Height          =   495
            Left            =   4680
            TabIndex        =   143
            Top             =   840
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   873
            _Version        =   327682
            LargeChange     =   10
            Min             =   1
            Max             =   100
            SelStart        =   50
            Value           =   45
         End
         Begin ComctlLib.ImageList ImageListback 
            Left            =   3000
            Top             =   360
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            MaskColor       =   12632256
            _Version        =   327682
         End
         Begin VB.Label lopnmapjpgtext 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2880
            TabIndex        =   140
            Top             =   1800
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.Label Label4 
            Caption         =   "BGM�G"
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
            Left            =   5400
            TabIndex        =   124
            Top             =   2280
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "BGM���q�G"
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
            Left            =   3720
            TabIndex        =   123
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "SE���q�G"
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
            Left            =   3960
            TabIndex        =   122
            Top             =   960
            Width           =   855
         End
         Begin VB.Label bgmve 
            Alignment       =   1  '�a�k���
            BackStyle       =   0  '�z��
            Caption         =   "50"
            Height          =   375
            Left            =   7320
            TabIndex        =   121
            Top             =   480
            Width           =   495
         End
         Begin VB.Label seve 
            Alignment       =   1  '�a�k���
            BackStyle       =   0  '�z��
            Caption         =   "45"
            Height          =   375
            Left            =   7320
            TabIndex        =   120
            Top             =   960
            Width           =   495
         End
         Begin VB.Label lopnmusictext 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   5400
            TabIndex        =   119
            Top             =   1440
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.Image ���a��� 
            Height          =   2415
            Left            =   240
            Stretch         =   -1  'True
            Top             =   360
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Label randomtext 
            Alignment       =   2  '�m�����
            BackStyle       =   0  '�z��
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   48
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   960
            TabIndex        =   118
            Top             =   960
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Shape randombk 
            BackColor       =   &H00000000&
            BackStyle       =   1  '���z��
            BorderStyle     =   0  '�z��
            Height          =   2415
            Left            =   240
            Top             =   360
            Width           =   2535
         End
      End
   End
   Begin TabDlg.SSTab t1 
      Height          =   6615
      Left            =   0
      TabIndex        =   134
      Top             =   1560
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   697
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�L�n������"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "�@��"
      TabPicture(0)   =   "Formsetting.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "�ƥ�d�]�ϥΪ̡^"
      TabPicture(1)   =   "Formsetting.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "�ƥ�d�]�q���^"
      TabPicture(2)   =   "Formsetting.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "�t��"
      TabPicture(3)   =   "Formsetting.frx":0D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
   End
   Begin VB.Frame �ƥ�d_�ϥΪ� 
      Caption         =   "�ƥ�d�s��(�ϥΪ̤�)"
      Height          =   5895
      Left            =   9600
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   9015
      Begin VB.PictureBox persontgrenus 
         Appearance      =   0  '����
         BorderStyle     =   0  '�S���ؽu
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   6000
         ScaleHeight     =   495
         ScaleWidth      =   2775
         TabIndex        =   98
         Top             =   240
         Width           =   2775
         Begin VB.OptionButton persontgruonus 
            Caption         =   "�H��"
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   4
            Left            =   1080
            TabIndex        =   102
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton persontgruonus 
            Caption         =   "��̤ܳj��"
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   1080
            TabIndex        =   101
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton persontgruonus 
            Caption         =   "�ۭq"
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   120
            TabIndex        =   100
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton persontgruonus 
            Caption         =   "�L"
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   120
            TabIndex        =   99
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.CheckBox persontgreus 
         Caption         =   "��uUnlight�ƥ�d�W�h"
         BeginProperty Font 
            Name            =   "�L�n������"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   97
         Top             =   360
         Value           =   1  '�֨�
         Width           =   2535
      End
      Begin VB.Frame f3 
         Height          =   1695
         Left            =   120
         TabIndex        =   36
         Top             =   4080
         Width           =   8535
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   13
            Left            =   2400
            TabIndex        =   42
            Top             =   240
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   14
            Left            =   2400
            TabIndex        =   41
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   15
            Left            =   2400
            TabIndex        =   40
            Top             =   1200
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   16
            Left            =   6000
            TabIndex        =   39
            Top             =   240
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   17
            Left            =   6000
            TabIndex        =   38
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   18
            Left            =   6000
            TabIndex        =   37
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Label persontgus 
            Caption         =   "1:"
            Height          =   375
            Index           =   13
            Left            =   2040
            TabIndex        =   50
            Top             =   240
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "2:"
            Height          =   375
            Index           =   14
            Left            =   2040
            TabIndex        =   49
            Top             =   720
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "3:"
            Height          =   375
            Index           =   15
            Left            =   2040
            TabIndex        =   48
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "4:"
            Height          =   375
            Index           =   16
            Left            =   5640
            TabIndex        =   47
            Top             =   240
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "5:"
            Height          =   375
            Index           =   17
            Left            =   5640
            TabIndex        =   46
            Top             =   720
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "6:"
            Height          =   375
            Index           =   18
            Left            =   5640
            TabIndex        =   45
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label personlevelus 
            BackStyle       =   0  '�z��
            Caption         =   "LV"
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
            Left            =   120
            TabIndex        =   44
            Top             =   720
            Width           =   495
         End
         Begin VB.Label personnameus 
            Caption         =   "XXX"
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   600
            TabIndex        =   43
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.Frame f2 
         Height          =   1695
         Left            =   120
         TabIndex        =   21
         Top             =   2400
         Width           =   8535
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   7
            Left            =   2400
            TabIndex        =   27
            Top             =   240
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   8
            Left            =   2400
            TabIndex        =   26
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   9
            Left            =   2400
            TabIndex        =   25
            Top             =   1200
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   10
            Left            =   6000
            TabIndex        =   24
            Top             =   240
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   11
            Left            =   6000
            TabIndex        =   23
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   12
            Left            =   6000
            TabIndex        =   22
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Label persontgus 
            Caption         =   "1:"
            Height          =   375
            Index           =   7
            Left            =   2040
            TabIndex        =   35
            Top             =   240
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "2:"
            Height          =   375
            Index           =   8
            Left            =   2040
            TabIndex        =   34
            Top             =   720
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "3:"
            Height          =   375
            Index           =   9
            Left            =   2040
            TabIndex        =   33
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "4:"
            Height          =   375
            Index           =   10
            Left            =   5640
            TabIndex        =   32
            Top             =   240
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "5:"
            Height          =   375
            Index           =   11
            Left            =   5640
            TabIndex        =   31
            Top             =   720
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "6:"
            Height          =   375
            Index           =   12
            Left            =   5640
            TabIndex        =   30
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label personlevelus 
            BackStyle       =   0  '�z��
            Caption         =   "LV"
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
            Left            =   120
            TabIndex        =   29
            Top             =   720
            Width           =   495
         End
         Begin VB.Label personnameus 
            Caption         =   "XXX"
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   600
            TabIndex        =   28
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.Frame f1 
         Height          =   1695
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   8535
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   6
            Left            =   6000
            TabIndex        =   13
            Top             =   1200
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   5
            Left            =   6000
            TabIndex        =   12
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   4
            Left            =   6000
            TabIndex        =   11
            Top             =   240
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   3
            Left            =   2400
            TabIndex        =   10
            Top             =   1200
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   2400
            TabIndex        =   9
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox personus 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   2400
            TabIndex        =   8
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label personnameus 
            Caption         =   "XXX"
            BeginProperty Font 
               Name            =   "�L�n������"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   600
            TabIndex        =   20
            Top             =   720
            Width           =   735
         End
         Begin VB.Label personlevelus 
            BackStyle       =   0  '�z��
            Caption         =   "LV"
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
            TabIndex        =   19
            Top             =   720
            Width           =   495
         End
         Begin VB.Label persontgus 
            Caption         =   "6:"
            Height          =   375
            Index           =   6
            Left            =   5640
            TabIndex        =   18
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "5:"
            Height          =   375
            Index           =   5
            Left            =   5640
            TabIndex        =   17
            Top             =   720
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "4:"
            Height          =   375
            Index           =   4
            Left            =   5640
            TabIndex        =   16
            Top             =   240
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "3:"
            Height          =   375
            Index           =   3
            Left            =   2040
            TabIndex        =   15
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "2:"
            Height          =   375
            Index           =   2
            Left            =   2040
            TabIndex        =   14
            Top             =   720
            Width           =   255
         End
         Begin VB.Label persontgus 
            Caption         =   "1:"
            Height          =   375
            Index           =   1
            Left            =   2040
            TabIndex        =   7
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Label personwagus 
         Caption         =   "(���⦳�H���ɵL�k�ۭq)"
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
         Left            =   3840
         TabIndex        =   110
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  '����
      BorderStyle     =   0  '�S���ؽu
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   9255
      TabIndex        =   3
      Top             =   8040
      Width           =   9255
      Begin VB.Label Label3 
         Alignment       =   2  '�m�����
         BackStyle       =   0  '�z��
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "�L�n������"
            Size            =   15.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4080
         TabIndex        =   4
         Top             =   120
         Width           =   975
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   3480
         Picture         =   "Formsetting.frx":0D3A
         Top             =   120
         Width           =   2145
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  '���z��
         Height          =   525
         Left            =   -120
         Top             =   120
         Width           =   9375
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BorderStyle     =   0  '�S���ؽu
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   240
      Picture         =   "Formsetting.frx":18BD
      ScaleHeight     =   1455
      ScaleWidth      =   1455
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '�z��
      Caption         =   "�j�p�j�A�o�̯�����z����ԲӪ��]�w��"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   5775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '�z��
      Caption         =   "�ٻݭn�����?"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   960
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '���z��
      Height          =   1845
      Left            =   0
      Top             =   -240
      Width           =   9255
   End
End
Attribute VB_Name = "Formsetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub checktest_Click()
If checktest.Value = 1 Then
    persontgreus.Value = 0
    persontgrecom.Value = 0
    persontgruoncom(3).Value = True
    persontgruonus(3).Value = True
    Formsetting.��Ԧa�Ͽ��.ListIndex = Formsetting.��Ԧa�Ͽ��.ListCount - 1
Else
    persontgreus.Value = 1
    persontgrecom.Value = 1
    persontgruoncom(1).Value = True
    persontgruonus(1).Value = True
    Formsetting.��Ԧa�Ͽ��.ListIndex = 0
End If
End Sub

Private Sub chkusesimilarlevel_Click()
If chkusesimilarlevel.Value = 0 Then
   cbsimilarlevel.Enabled = False
Else
    cbsimilarlevel.Enabled = True
End If
End Sub

Private Sub ckbgmmute_Click()
If ckbgmmute.Value = 1 Then
   FormMainMode.cMusicPlayer(0).Mute = True
Else
   FormMainMode.cMusicPlayer(0).Mute = False
End If
End Sub

Private Sub ckendturn_Click()
If ckendturn.Value = 1 Then
    ckendturnnum.Enabled = True
Else
    ckendturnnum.Enabled = False
End If
End Sub

Private Sub cksemute_Click()
Dim i As Integer
If cksemute.Value = 1 Then
    For i = 1 To FormMainMode.cMusicPlayer.UBound
        FormMainMode.cMusicPlayer(i).Mute = True
    Next
Else
    For i = 1 To FormMainMode.cMusicPlayer.UBound
        FormMainMode.cMusicPlayer(i).Mute = False
    Next
End If
End Sub

Private Sub Command1_Click()
On Error GoTo VBEError
cdgBGMchooce.ShowOpen
If lopnmusictext.Caption = "" And cdgBGMchooce.filename <> "" Then
    BGM���.AddItem "�m��L�n"
End If
If cdgBGMchooce.filename <> "" Then
    lopnmusictext.Caption = cdgBGMchooce.filename
    BGM���.ListIndex = BGM���.ListCount - 1
Else
    BGM���.ListIndex = 0
End If
Exit Sub
'===============
VBEError:
BGM���.ListIndex = 0
End Sub

Private Sub Command2_Click()
On Error GoTo VBEError
cdgMAPchooce.ShowOpen
If lopnmapjpgtext.Caption = "" And cdgMAPchooce.filename <> "" Then
    ��Ԧa�Ͽ��.AddItem "�m��L�n"
    ImageListback.ListImages.Add 15, , LoadPicture(cdgMAPchooce.filename)
ElseIf lopnmapjpgtext.Caption <> "" And cdgMAPchooce.filename <> "" Then
    ImageListback.ListImages.Remove 15
    ImageListback.ListImages.Add 15, , LoadPicture(cdgMAPchooce.filename)
End If
If cdgMAPchooce.filename <> "" Then
    lopnmapjpgtext.Caption = cdgMAPchooce.filename
    ��Ԧa�Ͽ��.ListIndex = ��Ԧa�Ͽ��.ListCount - 1
Else
    ��Ԧa�Ͽ��.ListIndex = 0
End If
Formsetting.��Ԧa�Ͽ��_Click
Exit Sub
'===============
VBEError:
��Ԧa�Ͽ��.ListIndex = 0
End Sub

Private Sub Form_Activate()
Dim i As Integer
Formsetting.Width = 9315
Formsetting.Height = 9150
For i = 1 To 3
   If FormMainMode.personnameus(i).Text <> "�m�H���n" Then
       personlevelus(i).Caption = FormMainMode.personlevelus(i).Text
       personnameus(i).Caption = FormMainMode.personnameus(i).Text
   Else
       personlevelus(i).Caption = ""
       personnameus(i).Caption = FormMainMode.personnameus(i).Text
   End If
   If FormMainMode.personnamecom(i).Text <> "�m�H���n" Then
       personlevelcom(i).Caption = FormMainMode.personlevelcom(i).Text
       personnamecom(i).Caption = FormMainMode.personnamecom(i).Text
   Else
       personlevelcom(i).Caption = ""
       personnamecom(i).Caption = FormMainMode.personnamecom(i).Text
   End If
Next
End Sub

Private Sub Form_Load()
Dim i As Integer

��Ԧa�Ͽ��.AddItem "�m�H���n"
��Ԧa�Ͽ��.AddItem "�H��Ӧa"
��Ԧa�Ͽ��.AddItem "���]�������۰}"
��Ԧa�Ͽ��.AddItem "�B�ʴ�`(�s)"
��Ԧa�Ͽ��.AddItem "�B�ʴ�`(��)"
��Ԧa�Ͽ��.AddItem "�U������"
��Ԧa�Ͽ��.AddItem "���ɯ"
��Ԧa�Ͽ��.AddItem "�Q�i�����´�"
��Ԧa�Ͽ��.AddItem "�ܤB���|�櫰��"
��Ԧa�Ͽ��.AddItem "�ƨg�s��"
��Ԧa�Ͽ��.AddItem "���Y����"
��Ԧa�Ͽ��.AddItem "���b�˪L"
��Ԧa�Ͽ��.AddItem "ÿ�e�઺���"
��Ԧa�Ͽ��.AddItem "�]�k�s��"
��Ԧa�Ͽ��.AddItem "�]��ù�e�����J"
'��Ԧa�Ͽ��.ListIndex = 0
'=============
For i = 1 To 14
    ImageListback.ListImages.Add i, , LoadPicture(app_path & "gif\system\map\" & i & ".jpg")
Next
'=============
BGM���.AddItem "�m�H��-�a�ϲզX�n"
BGM���.AddItem "�m�H���n"
BGM���.AddItem "�H��Ӧa"
'BGM���.AddItem "���]�������۰}"
BGM���.AddItem "�B�ʴ�`(�s)"
BGM���.AddItem "�U������"
BGM���.AddItem "���ɯ"
BGM���.AddItem "�Q�i�����´�"
BGM���.AddItem "�ܤB���|�櫰��"
BGM���.AddItem "�ƨg�s��"
BGM���.AddItem "���Y����"
BGM���.AddItem "���b�˪L"
BGM���.AddItem "ÿ�e�઺���"
BGM���.AddItem "�]��ù�e�����J"
BGM���.AddItem "�ª�"
'BGM���.ListIndex = 0
'=============
cdgBGMchooce.Filter = "MP3������(*.mp3)|*.mp3|Wave���T��(*.wav)|*.wav|MP4������(*.m4a)|*.m4a|�Ҧ��ɮ�(*.*)|*.*"
cdgMAPchooce.Filter = "JPG�Ϥ���(*.jpg)|*.jpg|BMP�I�}��(*.bmp)|*.bmp|�Ҧ��ɮ�(*.*)|*.*"
For i = 1 To 18
    personus(i).AddItem "(�L)"
    personus(i).AddItem "�C1"
    personus(i).AddItem "�C2"
    personus(i).AddItem "�C3"
    personus(i).AddItem "�C4"
    personus(i).AddItem "�C5"
    personus(i).AddItem "�C6"
    personus(i).AddItem "�C7"
    personus(i).AddItem "�C8"
    personus(i).AddItem "�j1"
    personus(i).AddItem "�j2"
    personus(i).AddItem "�j3"
    personus(i).AddItem "�j4"
    personus(i).AddItem "�j5"
    personus(i).AddItem "�j6"
    personus(i).AddItem "�j7"
    personus(i).AddItem "�j8"
    personus(i).AddItem "�S1"
    personus(i).AddItem "�S2"
    personus(i).AddItem "�S3"
    personus(i).AddItem "�S4"
    personus(i).AddItem "�S5"
    personus(i).AddItem "��1"
    personus(i).AddItem "��2"
    personus(i).AddItem "��3"
    personus(i).AddItem "��4"
    personus(i).AddItem "��5"
    personus(i).AddItem "��7"
    personus(i).AddItem "��1"
    personus(i).AddItem "��2"
    personus(i).AddItem "��3"
    personus(i).AddItem "��4"
    personus(i).AddItem "��5"
    personus(i).AddItem "���|1"
    personus(i).AddItem "���|2"
    personus(i).AddItem "���|3"
    personus(i).AddItem "���|4"
    personus(i).AddItem "���|5"
    personus(i).AddItem "�A�G�N1"
    personus(i).AddItem "�A�G�N2"
    personus(i).AddItem "�A�G�N3"
    personus(i).AddItem "�A�G�N5"
    personus(i).AddItem "HP�^�_1"
    personus(i).AddItem "HP�^�_2"
    personus(i).AddItem "HP�^�_3"
    personus(i).AddItem "�C3/�j1"
    personus(i).AddItem "�C4/�j2"
    personus(i).AddItem "�C5/�j3"
    personus(i).AddItem "�j3/�C1"
    personus(i).AddItem "�j4/�C2"
    personus(i).AddItem "�j5/�C3"
    personus(i).AddItem "��3/��1"
    personus(i).AddItem "��4/��1"
    personus(i).AddItem "��5/��1"
    personus(i).AddItem "�S1/��1"
    personus(i).AddItem "�S2/��2"
    personus(i).AddItem "�S3/��3"
    personus(i).AddItem "�C3/��1"
    personus(i).AddItem "�C4/��1"
    personus(i).AddItem "�C5/��1"
    personus(i).AddItem "�j3/��1"
    personus(i).AddItem "�j4/��1"
    personus(i).AddItem "�j5/��1"
    personus(i).AddItem "�C3/��1"
    personus(i).AddItem "�j3/��1"
    personus(i).AddItem "��1/�S1"
    personus(i).AddItem "��2/�S2"
    personus(i).AddItem "��3/�S3"
    personcom(i).AddItem "(�L)"
    personcom(i).AddItem "�C1"
    personcom(i).AddItem "�C2"
    personcom(i).AddItem "�C3"
    personcom(i).AddItem "�C4"
    personcom(i).AddItem "�C5"
    personcom(i).AddItem "�C6"
    personcom(i).AddItem "�C7"
    personcom(i).AddItem "�C8"
    personcom(i).AddItem "�j1"
    personcom(i).AddItem "�j2"
    personcom(i).AddItem "�j3"
    personcom(i).AddItem "�j4"
    personcom(i).AddItem "�j5"
    personcom(i).AddItem "�j6"
    personcom(i).AddItem "�j7"
    personcom(i).AddItem "�j8"
    personcom(i).AddItem "�S1"
    personcom(i).AddItem "�S2"
    personcom(i).AddItem "�S3"
    personcom(i).AddItem "�S4"
    personcom(i).AddItem "�S5"
    personcom(i).AddItem "��1"
    personcom(i).AddItem "��2"
    personcom(i).AddItem "��3"
    personcom(i).AddItem "��4"
    personcom(i).AddItem "��5"
    personcom(i).AddItem "��7"
    personcom(i).AddItem "��1"
    personcom(i).AddItem "��2"
    personcom(i).AddItem "��3"
    personcom(i).AddItem "��4"
    personcom(i).AddItem "��5"
    personcom(i).AddItem "���|1"
    personcom(i).AddItem "���|2"
    personcom(i).AddItem "���|3"
    personcom(i).AddItem "���|4"
    personcom(i).AddItem "���|5"
    personcom(i).AddItem "�A�G�N1"
    personcom(i).AddItem "�A�G�N2"
    personcom(i).AddItem "�A�G�N3"
    personcom(i).AddItem "�A�G�N5"
    personcom(i).AddItem "HP�^�_1"
    personcom(i).AddItem "HP�^�_2"
    personcom(i).AddItem "HP�^�_3"
    personcom(i).AddItem "�C3/�j1"
    personcom(i).AddItem "�C4/�j2"
    personcom(i).AddItem "�C5/�j3"
    personcom(i).AddItem "�j3/�C1"
    personcom(i).AddItem "�j4/�C2"
    personcom(i).AddItem "�j5/�C3"
    personcom(i).AddItem "��3/��1"
    personcom(i).AddItem "��4/��1"
    personcom(i).AddItem "��5/��1"
    personcom(i).AddItem "�S1/��1"
    personcom(i).AddItem "�S2/��2"
    personcom(i).AddItem "�S3/��3"
    personcom(i).AddItem "�C3/��1"
    personcom(i).AddItem "�C4/��1"
    personcom(i).AddItem "�C5/��1"
    personcom(i).AddItem "�j3/��1"
    personcom(i).AddItem "�j4/��1"
    personcom(i).AddItem "�j5/��1"
    personcom(i).AddItem "�C3/��1"
    personcom(i).AddItem "�j3/��1"
    personcom(i).AddItem "��1/�S1"
    personcom(i).AddItem "��2/�S2"
    personcom(i).AddItem "��3/�S3"
    persontgus(i).Visible = False
    persontgcom(i).Visible = False
Next
'checktest.Value = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = 1
Image2_Click
End Sub

Private Sub Image2_Click()
Formsetting.Visible = False
If Val(�D�ԼҦ��ﶵ_�P��.Text) > 30 Then �D�ԼҦ��ﶵ_�P��.Text = 30
If Val(�j�ð��Ҧ��ﶵ_�P��.Text) < 1 Then �j�ð��Ҧ��ﶵ_�P��.Text = 1
If Val(ckendturnnum.Text) <= 0 Then
    ckendturnnum.Text = 18
End If
End Sub

Private Sub Label3_Click()
Formsetting.Visible = False
If Val(�D�ԼҦ��ﶵ_�P��.Text) > 30 Then �D�ԼҦ��ﶵ_�P��.Text = 30
If Val(�j�ð��Ҧ��ﶵ_�P��.Text) < 1 Then �j�ð��Ҧ��ﶵ_�P��.Text = 1
If Val(ckendturnnum.Text) <= 0 Then
    ckendturnnum.Text = 18
End If
End Sub

Private Sub Label9_Click()

End Sub

Private Sub personcom1_Change(Index As Integer)

End Sub

Private Sub personcom1_Click(Index As Integer)

End Sub


Private Sub personcom_Click(Index As Integer)
If personcom(Index).ListIndex <> 0 Then
    If �@��t����.�ƥ�d��Ʈw(personcom(Index).Text, 1) <> persontgcom(Index).Caption And _
        �@��t����.�ƥ�d��Ʈw(personcom(Index).Text, 1) <> 0 And persontgrecom.Value = 1 Then
        MsgBox "���ƥ�d�H�Ϩ�����ϥέ�h!", 64, "UnlightVBE�t�δ���"
        personcom(Index).ListIndex = 0
    End If
End If
End Sub


Private Sub personnamecom_Change(Index As Integer)
If persontgrecom.Value = 1 Then
    If FormMainMode.opnpersonvs(2).Value = True Then
        If personnamecom(Index).Caption = "�m�H���n" Then
            If persontgruoncom(2).Value = True Then
                persontgruoncom(2).Value = False
                persontgruoncom(1).Value = True
                persontgruoncom_Click (1)
            End If
            persontgruoncom(2).Enabled = False
            personwagcom.Visible = True
        ElseIf personnamecom(1).Caption <> "�m�H���n" And personnamecom(2).Caption <> "�m�H���n" And _
            personnamecom(3).Caption <> "�m�H���n" Then
            persontgruoncom(2).Enabled = True
            personwagcom.Visible = False
        End If
    Else
        If personnamecom(1).Caption = "�m�H���n" Then
            If persontgruoncom(2).Value = True Then
                persontgruoncom(2).Value = False
                persontgruoncom(1).Value = True
                persontgruoncom_Click (1)
            End If
            persontgruoncom(2).Enabled = False
            personwagcom.Visible = True
        ElseIf personnamecom(1).Caption <> "�m�H���n" Then
            persontgruoncom(2).Enabled = True
            personwagcom.Visible = False
        End If
    End If
End If
End Sub

Private Sub personnameus_Change(Index As Integer)
If persontgreus.Value = 1 Then
    If FormMainMode.opnpersonvs(2).Value = True Then
        If personnameus(Index).Caption = "�m�H���n" Then
            If persontgruonus(2).Value = True Then
                persontgruonus(2).Value = False
                persontgruonus(1).Value = True
                persontgruonus_Click (1)
            End If
            persontgruonus(2).Enabled = False
            personwagus.Visible = True
        ElseIf personnameus(1).Caption <> "�m�H���n" And personnameus(2).Caption <> "�m�H���n" And _
            personnameus(3).Caption <> "�m�H���n" Then
            persontgruonus(2).Enabled = True
            personwagus.Visible = False
        End If
    Else
        If personnameus(1).Caption = "�m�H���n" Then
            If persontgruonus(2).Value = True Then
                persontgruonus(2).Value = False
                persontgruonus(1).Value = True
                persontgruonus_Click (1)
            End If
            persontgruonus(2).Enabled = False
            personwagus.Visible = True
        ElseIf personnameus(1).Caption <> "�m�H���n" Then
            persontgruonus(2).Enabled = True
            personwagus.Visible = False
        End If
    End If
End If
End Sub

Private Sub persontgcom_Change(Index As Integer)
Select Case persontgcom(Index).Caption
    Case 0
        personcom(Index).BackColor = RGB(192, 192, 192)
        personcom(Index).ForeColor = RGB(255, 255, 255)
    Case 1
        personcom(Index).BackColor = RGB(255, 0, 0)
        personcom(Index).ForeColor = RGB(255, 255, 255)
    Case 2
        personcom(Index).BackColor = RGB(0, 255, 0)
        personcom(Index).ForeColor = RGB(0, 0, 0)
    Case 3
        personcom(Index).BackColor = RGB(0, 0, 255)
        personcom(Index).ForeColor = RGB(255, 255, 255)
    Case 4
        personcom(Index).BackColor = RGB(255, 0, 255)
        personcom(Index).ForeColor = RGB(255, 255, 255)
    Case 5
        personcom(Index).BackColor = RGB(255, 255, 255)
        personcom(Index).ForeColor = RGB(0, 0, 0)
    Case 6
        personcom(Index).BackColor = RGB(0, 0, 0)
        personcom(Index).ForeColor = RGB(255, 255, 255)
    Case 7
        personcom(Index).BackColor = RGB(255, 255, 0)
        personcom(Index).ForeColor = RGB(0, 0, 0)
End Select
End Sub

Private Sub persontgrecom_Click()
Dim i As Integer

If persontgrecom.Value = 1 Then
    For i = 1 To 18
       personcom_Click (i)
    Next
    For i = 1 To 3
       personnamecom_Change (i)
    Next
Else
   personwagcom.Visible = False
   persontgruoncom(2).Enabled = True
End If

    
End Sub

Private Sub persontgreus_Click()
Dim i As Integer

If persontgreus.Value = 1 Then
    For i = 1 To 18
       personus_Click (i)
    Next
    For i = 1 To 3
       personnameus_Change (i)
    Next
Else
   personwagus.Visible = False
   persontgruonus(2).Enabled = True
End If
End Sub

Private Sub persontgruoncom_Click(Index As Integer)
Dim i As Integer

Select Case Index
    Case 1
       For i = 1 To 18
           personcom(i).Enabled = False
       Next
    Case 2
       For i = 1 To 18
           personcom(i).Enabled = True
       Next
    Case 3
       For i = 1 To 18
           personcom(i).Enabled = False
       Next
    Case 4
       For i = 1 To 18
           personcom(i).Enabled = False
       Next
End Select
End Sub

Private Sub persontgruonus_Click(Index As Integer)
Dim i As Integer

Select Case Index
    Case 1
       For i = 1 To 18
           personus(i).Enabled = False
       Next
    Case 2
       For i = 1 To 18
           personus(i).Enabled = True
       Next
    Case 3
       For i = 1 To 18
           personus(i).Enabled = False
       Next
    Case 4
       For i = 1 To 18
           personus(i).Enabled = False
       Next
End Select
End Sub

Private Sub persontgus_Change(Index As Integer)
Select Case persontgus(Index).Caption
    Case 0
        personus(Index).BackColor = RGB(192, 192, 192)
        personus(Index).ForeColor = RGB(255, 255, 255)
    Case 1
        personus(Index).BackColor = RGB(255, 0, 0)
        personus(Index).ForeColor = RGB(255, 255, 255)
    Case 2
        personus(Index).BackColor = RGB(0, 255, 0)
        personus(Index).ForeColor = RGB(0, 0, 0)
    Case 3
        personus(Index).BackColor = RGB(0, 0, 255)
        personus(Index).ForeColor = RGB(255, 255, 255)
    Case 4
        personus(Index).BackColor = RGB(255, 0, 255)
        personus(Index).ForeColor = RGB(255, 255, 255)
    Case 5
        personus(Index).BackColor = RGB(255, 255, 255)
        personus(Index).ForeColor = RGB(0, 0, 0)
    Case 6
        personus(Index).BackColor = RGB(0, 0, 0)
        personus(Index).ForeColor = RGB(255, 255, 255)
    Case 7
        personus(Index).BackColor = RGB(255, 255, 0)
        personus(Index).ForeColor = RGB(0, 0, 0)
End Select
End Sub

Private Sub personus_Click(Index As Integer)
If personus(Index).ListIndex <> 0 Then
    If �@��t����.�ƥ�d��Ʈw(personus(Index).Text, 1) <> persontgus(Index).Caption And _
        �@��t����.�ƥ�d��Ʈw(personus(Index).Text, 1) <> 0 And persontgreus.Value = 1 Then
        MsgBox "���ƥ�d�H�Ϩ�����ϥέ�h!", 64, "UnlightVBE�t�δ���"
        personus(Index).ListIndex = 0
    End If
End If
End Sub

Private Sub sdrbgm_Change()
bgmve.Caption = sdrbgm.Value
End Sub

Private Sub sdrbgm_Scroll()
bgmve.Caption = sdrbgm.Value
FormMainMode.cMusicPlayer(0).Volume = sdrbgm.Value
End Sub

Private Sub sdrse_Change()
seve.Caption = sdrse.Value
End Sub

Private Sub sdrse_Scroll()
Dim i As Integer

seve.Caption = sdrse.Value
For i = 1 To FormMainMode.cMusicPlayer.UBound
    FormMainMode.cMusicPlayer(i).Volume = sdrse.Value
Next
End Sub
Private Sub t1_Click(PreviousTab As Integer)
Select Case t1.Tab
     Case 0
           �@��]�w_���.Left = 120
           �@��]�w_���.Top = 2040
           �@��]�w_���.Visible = True
           �ƥ�d_�ϥΪ�.Visible = False
           �ƥ�d_�q��.Visible = False
           �@��]�w_���.ZOrder
     Case 1
           �ƥ�d_�ϥΪ�.Left = 120
           �ƥ�d_�ϥΪ�.Top = 2040
           �ƥ�d_�ϥΪ�.Visible = True
           �ƥ�d_�q��.Visible = False
           �@��]�w_���.Visible = False
           �ƥ�d_�ϥΪ�.ZOrder
     Case 2
           �ƥ�d_�q��.Left = 120
           �ƥ�d_�q��.Top = 2040
           �ƥ�d_�q��.Visible = True
           �ƥ�d_�ϥΪ�.Visible = False
           �@��]�w_���.Visible = False
           �ƥ�d_�q��.ZOrder
     Case 3
           ���_�t��.Left = 120
           ���_�t��.Top = 2040
           �ƥ�d_�q��.Visible = False
           �ƥ�d_�ϥΪ�.Visible = False
           �@��]�w_���.Visible = False
           ���_�t��.Visible = True
           ���_�t��.ZOrder
End Select
End Sub

Private Sub �j�ð��Ҧ��ﶵ_�P��_Change()
Dim i As Integer, j As Integer, k As Integer
j = 1
Do While j <= Len(�j�ð��Ҧ��ﶵ_�P��.Text)
   k = 0
      For k = 0 To 9
         If Asc(Mid(�j�ð��Ҧ��ﶵ_�P��.Text, j, 1)) = Asc(k) Then
             j = j + 1
             Exit For
         End If
      Next
      If k = 10 Then
         MsgBox "�j�p�j�A�п�J�Ʀr��...", 64
         �j�ð��Ҧ��ﶵ_�P��.Text = ""
         Exit Sub
      End If
Loop
End Sub

Private Sub �j�ð��ﶵ_Click()
If �j�ð��ﶵ.Value = 1 Then
    �j�ð��Ҧ��ﶵ_�P��.Enabled = True
Else
    �j�ð��Ҧ��ﶵ_�P��.Enabled = False
End If
End Sub

Private Sub �D�ԼҦ��ﶵ_Click()
If �D�ԼҦ��ﶵ.Value = 1 Then
    �D�ԼҦ��ﶵ_�P��.Enabled = True
Else
   �D�ԼҦ��ﶵ_�P��.Enabled = False
End If
End Sub

Private Sub �D�ԼҦ��ﶵ_�P��_Change()
Dim i As Integer, j As Integer, k As Integer
j = 1
Do While j <= Len(�D�ԼҦ��ﶵ_�P��.Text)
   k = 0
      For k = 0 To 9
         If Asc(Mid(�D�ԼҦ��ﶵ_�P��.Text, j, 1)) = Asc(k) Then
             j = j + 1
             Exit For
         End If
      Next
      If k = 10 Then
         MsgBox "�j�p�j�A�п�J�Ʀr��...", 64
         �D�ԼҦ��ﶵ_�P��.Text = ""
         Exit Sub
      End If
Loop

End Sub

Sub ��Ԧa�Ͽ��_Click()
Dim i As Integer

Select Case ��Ԧa�Ͽ��.Text
   Case "�B�ʴ�`(��)"
      BGM���.ListIndex = 13
   Case "�]�k�s��"
      BGM���.ListIndex = 7
   Case "���]�������۰}"
      BGM���.ListIndex = 7
   Case Else
      For i = 0 To BGM���.ListCount - 1
         BGM���.ListIndex = i
         If ��Ԧa�Ͽ��.Text = BGM���.Text Then
            Exit For
         End If
      Next
End Select
If ��Ԧa�Ͽ��.ListIndex > 0 Then
   ���a���.Picture = ImageListback.ListImages(��Ԧa�Ͽ��.ListIndex).Picture
   ���a���.Visible = True
   randomtext.Visible = False
   randombk.Visible = False
Else
   ���a���.Visible = False
   randomtext.Visible = True
   randombk.Visible = True
End If

End Sub
