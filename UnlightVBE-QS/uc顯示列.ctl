VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.UserControl ��ܦC 
   BackStyle       =   0  '�z��
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12720
   ClipBehavior    =   0  '�L
   HitBehavior     =   2  '�ϥ�ø�ϰϰ�
   ScaleHeight     =   3150
   ScaleWidth      =   12720
   Windowless      =   -1  'True
   Begin VB.Timer trmovehide 
      Enabled         =   0   'False
      Interval        =   120
      Left            =   8330
      Top             =   1200
   End
   Begin VB.Timer trmoveshow 
      Enabled         =   0   'False
      Interval        =   130
      Left            =   5160
      Top             =   1080
   End
   Begin VB.Label g2 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BackStyle       =   0  '�z��
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Noto Sans T Chinese Black"
         Size            =   36
         Charset         =   136
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   7590
      TabIndex        =   1
      Top             =   -120
      Width           =   1400
   End
   Begin VB.Label g1 
      Alignment       =   2  '�m�����
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BackStyle       =   0  '�z��
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Noto Sans T Chinese Black"
         Size            =   39.75
         Charset         =   136
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   2440
      TabIndex        =   0
      Top             =   0
      Width           =   1400
   End
   Begin ImageX.aicAlphaImage bn42 
      Height          =   1335
      Left            =   4320
      Top             =   1800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2355
      Image           =   "uc��ܦC.ctx":0000
      Scaler          =   3
   End
   Begin ImageX.aicAlphaImage bn4 
      Height          =   1335
      Left            =   10320
      Top             =   1800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2355
      Image           =   "uc��ܦC.ctx":0018
      Scaler          =   3
   End
   Begin VB.Image moverightjpg 
      Height          =   720
      Index           =   0
      Left            =   10080
      Picture         =   "uc��ܦC.ctx":0030
      Top             =   480
      Width           =   525
   End
   Begin VB.Image moveleftjpg 
      Height          =   720
      Index           =   0
      Left            =   2760
      Picture         =   "uc��ܦC.ctx":01F5
      Top             =   360
      Width           =   525
   End
   Begin ImageX.aicAlphaImage bn32 
      Height          =   1335
      Left            =   2760
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2355
      Image           =   "uc��ܦC.ctx":03B8
      Scaler          =   3
   End
   Begin ImageX.aicAlphaImage bn22 
      Height          =   1095
      Left            =   1800
      Top             =   1680
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1931
      Image           =   "uc��ܦC.ctx":03D0
      Scaler          =   3
   End
   Begin ImageX.aicAlphaImage bn12 
      Height          =   975
      Left            =   600
      Top             =   1680
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
      Image           =   "uc��ܦC.ctx":03E8
      Scaler          =   3
   End
   Begin ImageX.aicAlphaImage bn3 
      Height          =   1455
      Left            =   8880
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2566
      Image           =   "uc��ܦC.ctx":0400
      Scaler          =   3
   End
   Begin ImageX.aicAlphaImage bn2 
      Height          =   1215
      Left            =   7320
      Top             =   1920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2143
      Image           =   "uc��ܦC.ctx":0418
      Scaler          =   3
   End
   Begin ImageX.aicAlphaImage bn1 
      Height          =   1215
      Left            =   6120
      Top             =   1800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2143
      Image           =   "uc��ܦC.ctx":0430
      Scaler          =   3
   End
   Begin ImageX.aicAlphaImage image2 
      Height          =   1095
      Left            =   6960
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1931
      Image           =   "uc��ܦC.ctx":0448
      Scaler          =   3
   End
   Begin ImageX.aicAlphaImage image1 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1931
      Image           =   "uc��ܦC.ctx":0460
      Scaler          =   3
   End
   Begin ImageX.aicAlphaImage aie1 
      Height          =   1575
      Left            =   -120
      Top             =   0
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   2778
      Image           =   "uc��ܦC.ctx":0478
      Scaler          =   3
      Props           =   17
   End
End
Attribute VB_Name = "��ܦC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_smallimage As String
Dim m_smallimageus As String
Dim m_smallimagecom As String
Dim m_movetn As Boolean
Dim m_g1 As Integer
Dim m_g2 As Integer
Dim m_smallimageusleft As Integer
Dim m_smallimagecomleft As Integer
Dim m_g1v As Boolean
Dim m_g2v As Boolean
Dim m_bnc As Integer
Dim m_moveleftnum As Integer
Dim m_moveleftio As Integer
Dim m_moverightnum As Integer
Dim m_moverightio As Integer
Dim ���ʹϤ���ܼ�(1 To 2, 1 To 3) As Integer '������ܭp�ƾ��Ȯ��ܼ�(1.�ϥΪ�/2.�q��,1.�ثe��/2.��V-(1)�V��(2)�V�~/3.�ؼг̤j��)
Dim ���ʹϤ���ܧ�����(1 To 2) As Boolean '������ܭp�ƾ��O�_�w�����ܼ�(1.�ϥΪ�/2.�q��)
Dim trmovehidetime As Integer '������ܭp�ƾ��Ȯ��ܼ�
Dim m_moveleftrightc As Boolean
Dim m_smallimageuswidth As Integer
Dim m_smallimagecomwidth As Integer
Dim m_personvs As Integer

Public Property Get �H���԰��H��() As Integer
   �H���԰��H�� = m_personvs
End Property
Public Property Let �H���԰��H��(ByVal new_�H���԰��H�� As Integer)
   m_personvs = new_�H���԰��H��
   PropertyChanged "�H���԰��H��"
   Select Case Me.�H���԰��H��
       Case 1
            bn1.Left = 4200
            bn1.Top = 100
            bn2.Left = 5340
            bn2.Top = 100
            bn3.Left = 6480
            bn3.Top = 100
            bn12.Left = 4200
            bn12.Top = 100
            bn22.Left = 5340
            bn22.Top = 100
            bn32.Left = 6480
            bn32.Top = 100
            bn4.Visible = False
            bn42.Visible = False
       Case 3
            bn1.Left = 4080
            bn1.Top = 100
            bn2.Left = 4920
            bn2.Top = 100
            bn3.Left = 6600
            bn3.Top = 100
            bn4.Left = 5760
            bn4.Top = 100
            bn12.Left = 4080
            bn12.Top = 100
            bn22.Left = 4920
            bn22.Top = 100
            bn32.Left = 6600
            bn32.Top = 100
            bn42.Left = 5760
            bn42.Top = 100
            bn4.Visible = True
            bn42.Visible = True
   End Select
End Property
Public Property Get ��ܦC�Ϥ�() As String
   ��ܦC�Ϥ� = m_smallimage
End Property
Public Property Get �ϥΪ̤�p�H���Ϥ�width() As Integer
   �ϥΪ̤�p�H���Ϥ�width = m_smallimageuswidth
End Property
Public Property Get �q����p�H���Ϥ�width() As Integer
   �q����p�H���Ϥ�width = m_smallimagecomwidth
End Property
Public Property Get ���ʤ�V�Ϥ����() As Boolean
   ���ʤ�V�Ϥ���� = m_moveleftrightc
End Property
Public Property Get �ϥΪ̤貾�ʭ�() As Integer
   �ϥΪ̤貾�ʭ� = m_moveleftnum
End Property
Public Property Get �ϥΪ̤貾�ʤ��~() As Integer
   �ϥΪ̤貾�ʤ��~ = m_moveleftio
End Property
Public Property Get �q���貾�ʭ�() As Integer
   �q���貾�ʭ� = m_moverightnum
End Property
Public Property Get �q���貾�ʤ��~() As Integer
   �q���貾�ʤ��~ = m_moverightio
End Property
Public Property Get ���ʶ��q�����() As Boolean
   ���ʶ��q����� = m_movetn
End Property
Public Property Get goi1() As Integer
   goi1 = m_g1
End Property
Public Property Get goi2() As Integer
   goi2 = m_g2
End Property
Public Property Get ���ʶ��q��ܭ�() As Integer
   ���ʶ��q��ܭ� = m_bnc
End Property
Public Property Get �ϥΪ̤�p�H���Ϥ�left() As Integer
   �ϥΪ̤�p�H���Ϥ�left = m_smallimageusleft
End Property
Public Property Get �q����p�H���Ϥ�left() As Integer
   �q����p�H���Ϥ�left = m_smallimagecomleft
End Property
Public Property Get goi1���() As Boolean
   goi1��� = m_g1v
End Property
Public Property Get goi2���() As Boolean
   goi2��� = m_g2v
End Property
Public Property Get �ϥΪ̤�p�H���Ϥ�() As String
   �ϥΪ̤�p�H���Ϥ� = m_smallimageus
End Property
Public Property Let �ϥΪ̤�p�H���Ϥ�(ByVal New_�ϥΪ̤�p�H���Ϥ� As String)
   m_smallimageus = New_�ϥΪ̤�p�H���Ϥ�
   PropertyChanged "�ϥΪ̤�p�H���Ϥ�"
   If Me.�ϥΪ̤�p�H���Ϥ� <> "" Then
       image1.AutoSize = True
       image1.AutoRedraw = True
       image1.LoadImage_FromFile Me.�ϥΪ̤�p�H���Ϥ�
       image1.Top = 0
       image1.Left = 0
       Me.�ϥΪ̤�p�H���Ϥ�width = image1.Width
       image1.Mirror = aiMirrorNone
    End If
End Property
Public Property Get �q����p�H���Ϥ�() As String
   �q����p�H���Ϥ� = m_smallimagecom
End Property
Public Property Let �q����p�H���Ϥ�(ByVal New_�q����p�H���Ϥ� As String)
   m_smallimagecom = New_�q����p�H���Ϥ�
   PropertyChanged "�q����p�H���Ϥ�"
   If Me.�q����p�H���Ϥ� <> "" Then
       Image2.AutoSize = True
       Image2.AutoRedraw = True
       Image2.LoadImage_FromFile Me.�q����p�H���Ϥ�
       Image2.Top = 0
       Image2.Left = 7680
       Image2.Mirror = aiMirrorHorizontal
    End If
    Me.�q����p�H���Ϥ�width = Image2.Width
End Property
Public Property Let ��ܦC�Ϥ�(ByVal new_��ܦC�Ϥ� As String)
   m_smallimage = new_��ܦC�Ϥ�
   PropertyChanged "��ܦC�Ϥ�"
   If Me.��ܦC�Ϥ� <> "" Then
       aie1.AutoRedraw = True
       aie1.AutoSize = True
       aie1.LoadImage_FromFile Me.��ܦC�Ϥ�
       aie1.Left = 0
       aie1.Top = 0
   End If
End Property
Public Property Let goi2(ByVal newgoi2 As Integer)
   m_g2 = newgoi2
   PropertyChanged "goi2"
   g2.Caption = Me.goi2
End Property
Public Property Let goi1(ByVal newgoi1 As Integer)
   m_g1 = newgoi1
   PropertyChanged "goi1"
   g1.Caption = Me.goi1
End Property
Public Property Let goi1���(ByVal newgoi1v As Boolean)
   m_g1v = newgoi1v
   PropertyChanged "goi1���"
   If Me.goi1��� = False Then
       g1.Visible = False
    Else
       g1.Visible = True
        If g1.FontName = "Noto Sans T Chinese Black" Then
            g1.Top = -160
            g1.FontSize = 40
        Else
            g1.Top = 0
            g1.FontSize = 36
        End If
    End If
End Property
Public Property Let goi2���(ByVal newgoi2v As Boolean)
   m_g2v = newgoi2v
   PropertyChanged "goi2���"
   If Me.goi2��� = False Then
       g2.Visible = False
    Else
       g2.Visible = True
        If g2.FontName = "Noto Sans T Chinese Black" Then
            g2.Top = -160
            g2.FontSize = 40
        Else
            g2.Top = 0
            g2.FontSize = 36
        End If
    End If
End Property
Public Property Let �ϥΪ̤�p�H���Ϥ�left(ByVal new�ϥΪ̤�p�H���Ϥ�left As Integer)
    m_smallimageusleft = new�ϥΪ̤�p�H���Ϥ�left
   PropertyChanged "�ϥΪ̤�p�H���Ϥ�left"
   image1.Left = Me.�ϥΪ̤�p�H���Ϥ�left
End Property
Public Property Let �q����p�H���Ϥ�left(ByVal new�q����p�H���Ϥ�left As Integer)
    m_smallimagecomleft = new�q����p�H���Ϥ�left
   PropertyChanged "�q����p�H���Ϥ�left"
   Image2.Left = Me.�q����p�H���Ϥ�left
End Property

Public Property Let ���ʶ��q��ܭ�(ByVal new���ʶ��q��ܭ� As Integer)
    m_bnc = new���ʶ��q��ܭ�
   PropertyChanged "���ʶ��q��ܭ�"
   ���ʶ��q�����_���q
End Property
Public Property Let �ϥΪ̤貾�ʭ�(ByVal new�ϥΪ̤貾�ʭ� As Integer)
   m_moveleftnum = new�ϥΪ̤貾�ʭ�
   PropertyChanged "�ϥΪ̤貾�ʭ�"
   ���ʹϤ���ܼ�(1, 3) = Me.�ϥΪ̤貾�ʭ�
End Property
Public Property Let �ϥΪ̤貾�ʤ��~(ByVal new�ϥΪ̤貾�ʤ��~ As Integer)
   m_moveleftio = new�ϥΪ̤貾�ʤ��~
   PropertyChanged "�ϥΪ̤貾�ʤ��~"
   ���ʹϤ���ܼ�(1, 2) = Me.�ϥΪ̤貾�ʤ��~
End Property
Public Property Let �q���貾�ʭ�(ByVal new�q���貾�ʭ� As Integer)
   m_moverightnum = new�q���貾�ʭ�
   PropertyChanged "�q���貾�ʭ�"
   ���ʹϤ���ܼ�(2, 3) = Me.�q���貾�ʭ�
End Property
Public Property Let �q���貾�ʤ��~(ByVal new�q���貾�ʤ��~ As Integer)
   m_moverightio = new�q���貾�ʤ��~
   PropertyChanged "�q���貾�ʤ��~"
   ���ʹϤ���ܼ�(2, 2) = Me.�q���貾�ʤ��~
End Property
Public Property Let �ϥΪ̤�p�H���Ϥ�width(ByVal new�ϥΪ̤�p�H���Ϥ�width As Integer)
   m_smallimageuswidth = new�ϥΪ̤�p�H���Ϥ�width
   PropertyChanged "�ϥΪ̤�p�H���Ϥ�width"
End Property
Public Property Let �q����p�H���Ϥ�width(ByVal new�q����p�H���Ϥ�width As Integer)
   m_smallimagecomwidth = new�q����p�H���Ϥ�width
   PropertyChanged "�q����p�H���Ϥ�width"
End Property
Public Property Let ���ʤ�V�Ϥ����(ByVal new���ʤ�V�Ϥ���� As Boolean)
   Dim i As Integer
   
   m_moveleftrightc = new���ʤ�V�Ϥ����
   PropertyChanged "���ʤ�V�Ϥ����"
   If Me.���ʤ�V�Ϥ���� = True Then
         ���ʹϤ���ܼ�(1, 1) = 1
         ���ʹϤ���ܼ�(2, 1) = 1
         ���ʹϤ���ܧ�����(1) = False
         ���ʹϤ���ܧ�����(2) = False
         '=============================
         For i = 1 To ���ʹϤ���ܼ�(1, 3)
             Load moveleftjpg(i)
             moveleftjpg(i).Left = 2400 + i * 300
             moveleftjpg(i).Top = 120
         Next
         For i = 1 To ���ʹϤ���ܼ�(2, 3)
             Load moverightjpg(i)
             moverightjpg(i).Left = 8520 - i * 300
             moverightjpg(i).Top = 120
         Next
         '=============================
         trmovehidetime = 1
         trmoveshow.Enabled = True
    End If
End Property
Public Property Let ���ʶ��q�����(ByVal new���ʶ��q����� As Boolean)
    m_movetn = new���ʶ��q�����
   PropertyChanged "���ʶ��q�����"
   If Me.���ʶ��q����� = True Then
       bn1.Visible = True
       bn2.Visible = True
       bn3.Visible = True
       If Me.�H���԰��H�� = 3 Then
           bn4.Visible = True
       Else
           bn4.Visible = False
       End If
       Me.���ʶ��q��ܭ� = 0
    Else
       bn1.Visible = False
       bn2.Visible = False
       bn3.Visible = False
       bn4.Visible = False
       bn12.Visible = False
       bn22.Visible = False
       bn32.Visible = False
       bn42.Visible = False
    End If
End Property
Sub ���ʶ��q�����_���q()
   Select Case Me.���ʶ��q��ܭ�
      Case 0
            bn12.Visible = False
            bn22.Visible = False
            bn32.Visible = False
            bn42.Visible = False
      Case 1
            bn12.Visible = True
            bn22.Visible = False
            bn32.Visible = False
            bn42.Visible = False
      Case 2
            bn12.Visible = False
            bn22.Visible = True
            bn32.Visible = False
            bn42.Visible = False
      Case 3
            bn12.Visible = False
            bn22.Visible = False
            bn32.Visible = True
            bn42.Visible = False
      Case 4
            bn12.Visible = False
            bn22.Visible = False
            bn32.Visible = False
            bn42.Visible = True
   End Select
End Sub
Private Sub aie1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
���ʶ��q�����_���q
End Sub

Private Sub bn1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
bn12.Visible = True
End Sub

Private Sub bn12_Click(ByVal Button As Integer)
Me.���ʶ��q��ܭ� = 1
End Sub

Private Sub bn2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
bn22.Visible = True
End Sub

Private Sub bn22_Click(ByVal Button As Integer)
Me.���ʶ��q��ܭ� = 2
End Sub

Private Sub bn3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
bn32.Visible = True
End Sub

Private Sub bn32_Click(ByVal Button As Integer)
Me.���ʶ��q��ܭ� = 3
End Sub



Private Sub bn4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
bn42.Visible = True
End Sub

Private Sub bn42_Click(ByVal Button As Integer)
Me.���ʶ��q��ܭ� = 4
End Sub

Private Sub trmovehide_Timer()
Dim i As Integer

Select Case trmovehidetime
 Case 2
   If ���ʹϤ���ܼ�(1, 2) = 1 And ���ʹϤ���ܼ�(2, 2) = 2 Then
     If ���ʹϤ���ܼ�(1, 3) > 0 And ���ʹϤ���ܼ�(2, 3) > 0 Then
       If moveleftjpg(���ʹϤ���ܼ�(1, 3)).Visible = True And moverightjpg(���ʹϤ���ܼ�(2, 3)).Visible = True Then
          moveleftjpg(���ʹϤ���ܼ�(1, 3)).Visible = False
          moverightjpg(���ʹϤ���ܼ�(2, 3)).Visible = False
          ���ʹϤ���ܼ�(1, 3) = ���ʹϤ���ܼ�(1, 3) - 1
          ���ʹϤ���ܼ�(2, 3) = ���ʹϤ���ܼ�(2, 3) - 1
          Exit Sub
       End If
     End If
   ElseIf ���ʹϤ���ܼ�(1, 2) = 2 And ���ʹϤ���ܼ�(2, 2) = 1 Then
     If ���ʹϤ���ܼ�(1, 3) > 0 And ���ʹϤ���ܼ�(2, 3) > 0 Then
       If moveleftjpg(���ʹϤ���ܼ�(1, 3)).Visible = True And moverightjpg(���ʹϤ���ܼ�(2, 3)).Visible = True Then
          moveleftjpg(���ʹϤ���ܼ�(1, 3)).Visible = False
          moverightjpg(���ʹϤ���ܼ�(2, 3)).Visible = False
          ���ʹϤ���ܼ�(1, 3) = ���ʹϤ���ܼ�(1, 3) - 1
          ���ʹϤ���ܼ�(2, 3) = ���ʹϤ���ܼ�(2, 3) - 1
          Exit Sub
       End If
     End If
   End If
   trmovehidetime = trmovehidetime + 1
 Case 10
      '===���p�����ŦX����ɤ��ʧ@
      trmovehide.Enabled = False
      '=========================
      For i = 1 To moveleftjpg.UBound
          moveleftjpg(i).Visible = False
          Unload moveleftjpg(i)
      Next
      For i = 1 To moverightjpg.UBound
          moverightjpg(i).Visible = False
          Unload moverightjpg(i)
      Next
      '=========================
      Me.���ʤ�V�Ϥ���� = False
 Case Else
      trmovehidetime = trmovehidetime + 1
End Select
End Sub

Private Sub trmoveshow_Timer()
If ���ʹϤ���ܼ�(1, 1) <= ���ʹϤ���ܼ�(1, 3) Then
   If ���ʹϤ���ܼ�(1, 2) = 1 Then
      moveleftjpg(���ʹϤ���ܼ�(1, 1)).Picture = LoadPicture(App.Path & "\gif\system\movein.gif")
      moveleftjpg(���ʹϤ���ܼ�(1, 1)).Visible = True
      moveleftjpg(���ʹϤ���ܼ�(1, 1)).ZOrder
   Else
      moveleftjpg(���ʹϤ���ܼ�(1, 1)).Picture = LoadPicture(App.Path & "\gif\system\moveout.gif")
      moveleftjpg(���ʹϤ���ܼ�(1, 1)).Visible = True
      moveleftjpg(���ʹϤ���ܼ�(1, 1)).ZOrder
   End If
   ���ʹϤ���ܼ�(1, 1) = ���ʹϤ���ܼ�(1, 1) + 1
Else
   ���ʹϤ���ܧ�����(1) = True
End If

If ���ʹϤ���ܼ�(2, 1) <= ���ʹϤ���ܼ�(2, 3) Then
   If ���ʹϤ���ܼ�(2, 2) = 1 Then
      moverightjpg(���ʹϤ���ܼ�(2, 1)).Picture = LoadPicture(App.Path & "\gif\system\moveout.gif")
      moverightjpg(���ʹϤ���ܼ�(2, 1)).Visible = True
      moverightjpg(���ʹϤ���ܼ�(2, 1)).ZOrder
   Else
      moverightjpg(���ʹϤ���ܼ�(2, 1)).Picture = LoadPicture(App.Path & "\gif\system\movein.gif")
      moverightjpg(���ʹϤ���ܼ�(2, 1)).Visible = True
      moverightjpg(���ʹϤ���ܼ�(2, 1)).ZOrder
   End If
   ���ʹϤ���ܼ�(2, 1) = ���ʹϤ���ܼ�(2, 1) + 1
Else
   ���ʹϤ���ܧ�����(2) = True
End If

If ���ʹϤ���ܧ�����(1) = True And ���ʹϤ���ܧ�����(2) = True Then
trmoveshow.Enabled = False
���ʹϤ���ܧ�����(1) = False
���ʹϤ���ܧ�����(2) = False
trmovehide.Enabled = True
End If
End Sub

Private Sub UserControl_Show()
bn1.AutoRedraw = True
bn1.AutoSize = True
bn2.AutoRedraw = True
bn2.AutoSize = True
bn3.AutoRedraw = True
bn3.AutoSize = True
bn4.AutoRedraw = True
bn4.AutoSize = True
bn12.AutoRedraw = True
bn12.AutoSize = True
bn22.AutoRedraw = True
bn22.AutoSize = True
bn32.AutoRedraw = True
bn32.AutoSize = True
bn42.AutoRedraw = True
bn42.AutoSize = True
bn1.LoadImage_FromFile App.Path & "\gif\system\left_1.png"
bn2.LoadImage_FromFile App.Path & "\gif\system\rest_1.png"
bn3.LoadImage_FromFile App.Path & "\gif\system\right_1.png"
bn4.LoadImage_FromFile App.Path & "\gif\system\change_1.png"
bn12.LoadImage_FromFile App.Path & "\gif\system\left_2.png"
bn22.LoadImage_FromFile App.Path & "\gif\system\rest_2.png"
bn32.LoadImage_FromFile App.Path & "\gif\system\right_2.png"
bn42.LoadImage_FromFile App.Path & "\gif\system\change_2.png"
Me.���ʶ��q����� = False
Me.���ʶ��q��ܭ� = 0

moveleftjpg(0).Left = 2400
moveleftjpg(0).Top = 120
moverightjpg(0).Left = 8520
moverightjpg(0).Top = 120
moveleftjpg(0).Visible = False
moverightjpg(0).Visible = False
Me.���ʤ�V�Ϥ���� = False
End Sub


