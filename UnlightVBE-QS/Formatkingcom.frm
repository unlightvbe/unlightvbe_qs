VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.Form Formatkingcom 
   BackColor       =   &H00000000&
   BorderStyle     =   1  '虫絬㏕﹚
   ClientHeight    =   9180
   ClientLeft      =   9480
   ClientTop       =   1965
   ClientWidth     =   5820
   Icon            =   "Formatkingcom.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   5820
   Begin VB.Timer t1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   8040
   End
   Begin VB.PictureBox atkingcomjpg_old 
      Appearance      =   0  'キ
      BackColor       =   &H00000000&
      BorderStyle     =   0  '⊿Τ絬
      ForeColor       =   &H80000008&
      Height          =   11025
      Left            =   8400
      Picture         =   "Formatkingcom.frx":0CCA
      ScaleHeight     =   11025
      ScaleWidth      =   18600
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   18600
   End
   Begin ImageX.aicAlphaImage atkingcomjpg 
      Height          =   9195
      Left            =   0
      Top             =   0
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   16219
      Image           =   "Formatkingcom.frx":15D3A
      Scaler          =   3
      Props           =   13
   End
End
Attribute VB_Name = "Formatkingcom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
╰参摸.瞒秨笴栏矗ボ Cancel, UnloadMode
End Sub

Private Sub t1_Timer()
If ヘ玡计(31) = 19 Then
   t1.Enabled = False
   Vss_AtkingStartPlayNum(3) = 1
   Unload Me
'==========================
ElseIf ヘ玡计(31) = 10 Then
   Vss_AtkingStartPlayNum(1) = 1 'м磅︽い传瓜
   Vss_AtkingStartPlayNum(2) = 1 'м磅︽い币笆
   ヘ玡计(31) = Val(ヘ玡计(31)) + 1
ElseIf ヘ玡计(31) = 7 Then
   ╰参摸.冀 5
   ヘ玡计(31) = Val(ヘ玡计(31)) + 1
ElseIf ヘ玡计(31) = 5 Then
   atkingcomjpg.Visible = True
   ヘ玡计(31) = Val(ヘ玡计(31)) + 1
Else
   ヘ玡计(31) = Val(ヘ玡计(31)) + 1
End If
DoEvents
End Sub
