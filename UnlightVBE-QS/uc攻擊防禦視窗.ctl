VERSION 5.00
Begin VB.UserControl uc§ðÀ»¨¾¿mµøµ¡ 
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  '³z©ú
   ClientHeight    =   8175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4515
   ClipBehavior    =   0  'µL
   DataBindingBehavior=   2  'vbComplexBound
   DataSourceBehavior=   1  'vbDataSource
   FillStyle       =   0  '¹ê¤ß
   BeginProperty Font 
      Name            =   "·L³n¥¿¶ÂÅé"
      Size            =   9
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   8175
   ScaleWidth      =   4515
   Begin VB.Timer trhide 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2760
      Top             =   3840
   End
   Begin VB.Timer t1 
      Enabled         =   0   'False
      Interval        =   130
      Left            =   2040
      Top             =   3840
   End
   Begin VB.Image adcom 
      Height          =   750
      Index           =   20
      Left            =   3480
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":0000
      Top             =   0
      Width           =   750
   End
   Begin VB.Image adcom 
      Height          =   750
      Index           =   19
      Left            =   2640
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":0690
      Top             =   0
      Width           =   750
   End
   Begin VB.Image adcom 
      Height          =   750
      Index           =   18
      Left            =   1800
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":0D20
      Top             =   0
      Width           =   750
   End
   Begin VB.Image adcom 
      Height          =   750
      Index           =   17
      Left            =   960
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":13B0
      Top             =   0
      Width           =   750
   End
   Begin VB.Image adcom 
      Height          =   750
      Index           =   16
      Left            =   120
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":1A40
      Top             =   0
      Width           =   750
   End
   Begin VB.Image adcom 
      Height          =   750
      Index           =   15
      Left            =   3480
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":20D0
      Top             =   840
      Width           =   750
   End
   Begin VB.Image adcom 
      Height          =   750
      Index           =   14
      Left            =   2640
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":2760
      Top             =   840
      Width           =   750
   End
   Begin VB.Image adcom 
      Height          =   750
      Index           =   13
      Left            =   1800
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":2DF0
      Top             =   840
      Width           =   750
   End
   Begin VB.Image adcom 
      Height          =   750
      Index           =   12
      Left            =   960
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":3480
      Top             =   840
      Width           =   750
   End
   Begin VB.Image adcom 
      Height          =   750
      Index           =   11
      Left            =   120
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":3B10
      Top             =   840
      Width           =   750
   End
   Begin VB.Image adcom 
      Height          =   750
      Index           =   10
      Left            =   3480
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":41A0
      Top             =   1680
      Width           =   750
   End
   Begin VB.Image adcom 
      Height          =   750
      Index           =   9
      Left            =   2640
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":4830
      Top             =   1680
      Width           =   750
   End
   Begin VB.Image adcom 
      Height          =   750
      Index           =   8
      Left            =   1800
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":4EC0
      Top             =   1680
      Width           =   750
   End
   Begin VB.Image adcom 
      Height          =   750
      Index           =   7
      Left            =   960
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":5550
      Top             =   1680
      Width           =   750
   End
   Begin VB.Image adcom 
      Height          =   750
      Index           =   6
      Left            =   120
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":5BE0
      Top             =   1680
      Width           =   750
   End
   Begin VB.Image adcom 
      Height          =   750
      Index           =   5
      Left            =   3480
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":6270
      Top             =   2520
      Width           =   750
   End
   Begin VB.Image adcom 
      Height          =   750
      Index           =   4
      Left            =   2640
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":6900
      Top             =   2520
      Width           =   750
   End
   Begin VB.Image adcom 
      Height          =   750
      Index           =   3
      Left            =   1800
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":6F90
      Top             =   2520
      Width           =   750
   End
   Begin VB.Image adcom 
      Height          =   750
      Index           =   2
      Left            =   960
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":7620
      Top             =   2520
      Width           =   750
   End
   Begin VB.Image adcom 
      Height          =   750
      Index           =   1
      Left            =   120
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":7CB0
      Top             =   2520
      Width           =   750
   End
   Begin VB.Image adus 
      Height          =   750
      Index           =   5
      Left            =   3480
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":8340
      Top             =   4680
      Width           =   750
   End
   Begin VB.Image adus 
      Height          =   750
      Index           =   4
      Left            =   2640
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":8632
      Top             =   4680
      Width           =   750
   End
   Begin VB.Image adus 
      Height          =   750
      Index           =   3
      Left            =   1800
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":8924
      Top             =   4680
      Width           =   750
   End
   Begin VB.Image adus 
      Height          =   750
      Index           =   2
      Left            =   960
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":8C16
      Top             =   4680
      Width           =   750
   End
   Begin VB.Image adus 
      Height          =   750
      Index           =   1
      Left            =   120
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":8F08
      Top             =   4680
      Width           =   750
   End
   Begin VB.Image adus 
      Height          =   750
      Index           =   10
      Left            =   3480
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":91FA
      Top             =   5520
      Width           =   750
   End
   Begin VB.Image adus 
      Height          =   750
      Index           =   9
      Left            =   2640
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":94EC
      Top             =   5520
      Width           =   750
   End
   Begin VB.Image adus 
      Height          =   750
      Index           =   8
      Left            =   1800
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":97DE
      Top             =   5520
      Width           =   750
   End
   Begin VB.Image adus 
      Height          =   750
      Index           =   7
      Left            =   960
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":9AD0
      Top             =   5520
      Width           =   750
   End
   Begin VB.Image adus 
      Height          =   750
      Index           =   6
      Left            =   120
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":9DC2
      Top             =   5520
      Width           =   750
   End
   Begin VB.Image adus 
      Height          =   750
      Index           =   15
      Left            =   3480
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":A0B4
      Top             =   6360
      Width           =   750
   End
   Begin VB.Image adus 
      Height          =   750
      Index           =   14
      Left            =   2640
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":A3A6
      Top             =   6360
      Width           =   750
   End
   Begin VB.Image adus 
      Height          =   750
      Index           =   13
      Left            =   1800
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":A698
      Top             =   6360
      Width           =   750
   End
   Begin VB.Image adus 
      Height          =   750
      Index           =   12
      Left            =   960
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":A98A
      Top             =   6360
      Width           =   750
   End
   Begin VB.Image adus 
      Height          =   750
      Index           =   11
      Left            =   120
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":AC7C
      Top             =   6360
      Width           =   750
   End
   Begin VB.Image adus 
      Height          =   750
      Index           =   20
      Left            =   3480
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":AF6E
      Top             =   7200
      Width           =   750
   End
   Begin VB.Image adus 
      Height          =   750
      Index           =   19
      Left            =   2640
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":B260
      Top             =   7200
      Width           =   750
   End
   Begin VB.Image adus 
      Height          =   750
      Index           =   18
      Left            =   1800
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":B552
      Top             =   7200
      Width           =   750
   End
   Begin VB.Image adus 
      Height          =   750
      Index           =   17
      Left            =   960
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":B844
      Top             =   7200
      Width           =   750
   End
   Begin VB.Image adus 
      Height          =   750
      Index           =   16
      Left            =   120
      Picture         =   "uc§ðÀ»¨¾¿mµøµ¡.ctx":BB36
      Top             =   7200
      Width           =   750
   End
End
Attribute VB_Name = "uc§ðÀ»¨¾¿mµøµ¡"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_adcomtext As Integer
Dim m_adustext As Integer
Dim adusall As Integer
Dim adcomall As Integer
Dim hideall As Integer
Dim m_adcgetext As Integer
Dim m_adwaittext As Boolean
 
 
Private Sub t1_Timer()
Select Case Me.adcge
Case 1
    If adusall <= Me.adust Then
       adus(adusall).Picture = LoadPicture(App.Path & "\gif\system\atkshow.gif")
       adus(adusall).Visible = True
       adusall = adusall + 1
    End If
    If adcomall <= Me.adcomt Then
       adcom(adcomall).Picture = LoadPicture(App.Path & "\gif\system\defshow.gif")
       adcom(adcomall).Visible = True
       adcomall = adcomall + 1
    End If
Case 2
   If adusall <= Me.adust Then
       adus(adusall).Picture = LoadPicture(App.Path & "\gif\system\defshow.gif")
       adus(adusall).Visible = True
       adusall = adusall + 1
    End If
    If adcomall <= Me.adcomt Then
       adcom(adcomall).Picture = LoadPicture(App.Path & "\gif\system\atkshow.gif")
       adcom(adcomall).Visible = True
       adcomall = adcomall + 1
    End If
End Select
If adusall > Me.adust And adcomall > Me.adcomt Then
   t1.Enabled = False
   hideall = 1
   trhide.Enabled = True
End If
End Sub

Private Sub trhide_Timer()
If hideall <= 20 Then
   If adus(hideall).Visible = True And adcom(hideall).Visible = True Then
      adus(hideall).Visible = False
      adcom(hideall).Visible = False
      hideall = hideall + 1
   Else
      trhide.Enabled = False
      Me.adwait = True
   End If
Else
   trhide.Enabled = False
   Me.adwait = True
End If
End Sub

Private Sub UserControl_Show()
Dim i As Integer
adusall = 1
adcomall = 1
t1.Enabled = True
For i = 1 To 20
   adus(i).Visible = False
   adcom(i).Visible = False
Next
End Sub

Public Property Get adcomt() As Integer
   adcomt = m_adcomtext
End Property

Public Property Get adust() As Integer
   adust = m_adustext
End Property
Public Property Get adcge() As Integer
   adcge = m_adcgetext
End Property
Public Property Get adwait() As Boolean
   adwait = m_adwaittext
End Property
Public Property Let adcomt(ByVal New_adcom As Integer)
   m_adcomtext = New_adcom
   PropertyChanged "adcomt"
End Property
Public Property Let adust(ByVal New_adus As Integer)
   m_adustext = New_adus
   PropertyChanged "adust"
End Property
Public Property Let adcge(ByVal New_adcge As Integer)
   m_adcgetext = New_adcge
   PropertyChanged "adcge"
End Property
Public Property Let adwait(ByVal New_adwait As Boolean)
   m_adwaittext = New_adwait
   PropertyChanged "adwait"
End Property
