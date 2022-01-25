VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.Form Formatkingus 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  '虫絬㏕﹚
   ClientHeight    =   9180
   ClientLeft      =   5085
   ClientTop       =   1275
   ClientWidth     =   11955
   ForeColor       =   &H00000000&
   Icon            =   "Formatkingus.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   11955
   Begin VB.Timer t1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   8160
   End
   Begin VB.PictureBox atkingusjpg_old 
      Appearance      =   0  'キ
      BackColor       =   &H00000000&
      BorderStyle     =   0  '⊿Τ絬
      ForeColor       =   &H80000008&
      Height          =   11025
      Left            =   8040
      Picture         =   "Formatkingus.frx":0CCA
      ScaleHeight     =   11025
      ScaleWidth      =   10680
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   10680
   End
   Begin ImageX.aicAlphaImage atkingusjpg 
      Height          =   9195
      Left            =   0
      Top             =   0
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   16219
      Image           =   "Formatkingus.frx":24C96
      Scaler          =   3
      Props           =   13
   End
End
Attribute VB_Name = "Formatkingus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
   YesNo = MsgBox("絋﹚瞒秨笴栏?", 36, "UnlightVBE-╰参矗ボ")
   If YesNo = 6 Then
    End
   Else
    Cancel = 1
   End If
End If
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
   atkingusjpg.Visible = True
   ヘ玡计(31) = Val(ヘ玡计(31)) + 1
Else
   ヘ玡计(31) = Val(ヘ玡计(31)) + 1
End If
DoEvents
End Sub
