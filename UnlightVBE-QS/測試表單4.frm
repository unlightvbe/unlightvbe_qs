VERSION 5.00
Begin VB.Form 測試表單4 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   BorderStyle     =   1  '單線固定
   Caption         =   "Form11"
   ClientHeight    =   9915
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "測試表單4.frx":0000
   ScaleHeight     =   9915
   ScaleWidth      =   11340
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   9840
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   8760
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin UnlightVBE.uc擲骰介面 uc擲骰介面1 
      Height          =   9910
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   17489
   End
   Begin UnlightVBE.ucCard ucCard1 
      Height          =   1335
      Index           =   1
      Left            =   2760
      TabIndex        =   5
      Top             =   2040
      Width           =   855
      _ExtentX        =   2143
      _ExtentY        =   2990
   End
   Begin UnlightVBE.ucCard ucCard1 
      Height          =   1335
      Index           =   0
      Left            =   1920
      TabIndex        =   4
      Top             =   2040
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   2355
   End
   Begin UnlightVBE.uc戰鬥系統牌型介面 uc戰鬥系統牌型介面1 
      Height          =   9915
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   17489
   End
End
Attribute VB_Name = "測試表單4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
uc戰鬥系統牌型介面1.Message = "決定攻擊力30點"
uc擲骰介面1.DiceInputMode = 2
uc擲骰介面1.diceusTotal = 30
uc擲骰介面1.dicecomTotal = 10
uc擲骰介面1.DiceATKMode = 1
uc擲骰介面1.PersonImage = App.Path & "\gif\Ayn\Aynperson1.png"
uc擲骰介面1.dicevoice = 40
uc擲骰介面1.DiceStart = True
End Sub

Private Sub Command2_Click()
uc戰鬥系統牌型介面1.Message = "決定攻擊力79點"
uc擲骰介面1.DiceInputMode = 2
uc擲骰介面1.diceusTotal = 38
uc擲骰介面1.dicecomTotal = 79
uc擲骰介面1.DiceATKMode = 2
uc擲骰介面1.PersonImage = App.Path & "\gif\Sheri\sheriperson2.png"
uc擲骰介面1.dicevoice = 40
uc擲骰介面1.DiceStart = True
End Sub

Private Sub Form_Load()
ucCard1(0).CardImage = App.Path & "\card\014-1.bmp"
ucCard1(1).CardImage = App.Path & "\card\014-1.bmp"
ucCard1(0).LocationType = 1
ucCard1(1).LocationType = 2
ucCard1(0).CardEventType = True
ucCard1(1).CardEventType = True
End Sub

Private Sub ucCard1_CardButtonClickout(Index As Integer)
'ucCard1(Index).LocationType = 3
End Sub

Private Sub ucCard1_CardMouseMove(Index As Integer)
'MsgBox 123456789
End Sub
