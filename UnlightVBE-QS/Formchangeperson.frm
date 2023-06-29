VERSION 5.00
Begin VB.Form Formchangeperson 
   BackColor       =   &H00000000&
   BorderStyle     =   1  '單線固定
   Caption         =   "交換角色"
   ClientHeight    =   4845
   ClientLeft      =   6690
   ClientTop       =   2535
   ClientWidth     =   6450
   BeginProperty Font 
      Name            =   "微軟正黑體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Formchangeperson.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6450
   Begin VB.Timer 使用者方智慧型AI_自動控制選人 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6120
      Top             =   4320
   End
   Begin VB.Timer trchange 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5880
      Top             =   360
   End
   Begin UnlightVBE.uc角色卡片介面 card 
      Height          =   3615
      Index           =   2
      Left            =   3480
      TabIndex        =   1
      Top             =   480
      Width           =   2535
      _ExtentX        =   2355
      _ExtentY        =   3625
   End
   Begin UnlightVBE.uc角色卡片介面 card 
      Height          =   3615
      Index           =   1
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2535
      _ExtentX        =   2355
      _ExtentY        =   3625
   End
   Begin VB.Image bnok 
      Height          =   345
      Index           =   2
      Left            =   3600
      Picture         =   "Formchangeperson.frx":0CCA
      Top             =   4200
      Width           =   2250
   End
   Begin VB.Image bnok 
      Height          =   345
      Index           =   1
      Left            =   480
      Picture         =   "Formchangeperson.frx":35A8
      Top             =   4200
      Width           =   2250
   End
End
Attribute VB_Name = "Formchangeperson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim changepersonComplete As Boolean
Sub bnok_Click(Index As Integer)
    If liveus(角色待機人物紀錄數(1, Index + 1)) > 0 Then
        Me.Hide
        changepersonComplete = True
        戰鬥系統類.人物交換_使用者_指定交換 Index + 1
        執行動作_交換人物角色_結束執行
        Unload Me
    End If
End Sub

Private Sub bnok_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    bnok(Index).Picture = LoadPicture(App.Path & "\gif\system\changeok_2.bmp")
End Sub

Private Sub Form_Load()
    changepersonComplete = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bnok(1).Picture = LoadPicture(App.Path & "\gif\system\changeok_1.bmp")
    bnok(2).Picture = LoadPicture(App.Path & "\gif\system\changeok_1.bmp")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim m As Integer
    Me.Hide
    If changepersonComplete = False Then
        If liveus(角色人物對戰人數(1, 2)) > 0 Then
            執行動作_交換人物角色_結束執行
        Else
            Randomize
            m = Int(Rnd() * 2) + 1
            If liveus(角色待機人物紀錄數(1, m + 1)) > 0 Then
                戰鬥系統類.人物交換_使用者_指定交換 m + 1
                執行動作_交換人物角色_結束執行
            Else
                If m = 1 Then m = 2 Else m = 1
                戰鬥系統類.人物交換_使用者_指定交換 m + 1
                執行動作_交換人物角色_結束執行
            End If
        End If
    End If
    Unload Me
End Sub

Sub 使用者方智慧型AI_自動控制選人_Timer()
    Dim i As Integer
    For i = 1 To 2
        If liveus(角色待機人物紀錄數(1, i + 1)) > 0 Then
            Formchangeperson.bnok_Click (i)
            Formchangeperson.使用者方智慧型AI_自動控制選人.Enabled = False
            Exit Sub
        End If
    Next
End Sub
