VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form 測試表單3 
   Caption         =   "Form10"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10080
   LinkTopic       =   "Form10"
   ScaleHeight     =   7785
   ScaleWidth      =   10080
   StartUpPosition =   3  '系統預設值
   Begin VB.PictureBox card 
      BorderStyle     =   0  '沒有框線
      Height          =   1335
      Index           =   0
      Left            =   4440
      Picture         =   "測試表單3.frx":0000
      ScaleHeight     =   1335
      ScaleWidth      =   855
      TabIndex        =   7
      Top             =   1560
      Width           =   855
      Begin VB.Image cgen 
         Height          =   330
         Index           =   0
         Left            =   480
         Picture         =   "測試表單3.frx":4722
         Top             =   240
         Width           =   330
      End
      Begin VB.Image cge 
         Height          =   330
         Index           =   0
         Left            =   480
         Picture         =   "測試表單3.frx":4B43
         Top             =   600
         Width           =   330
      End
      Begin VB.Image cqe 
         Height          =   330
         Index           =   0
         Left            =   0
         Picture         =   "測試表單3.frx":4DDD
         Top             =   240
         Width           =   330
      End
      Begin VB.Image cqen 
         Height          =   330
         Index           =   0
         Left            =   0
         Picture         =   "測試表單3.frx":5075
         Top             =   600
         Width           =   330
      End
      Begin VB.Image cgu 
         Height          =   225
         Index           =   0
         Left            =   240
         Picture         =   "測試表單3.frx":549D
         Top             =   0
         Width           =   300
      End
      Begin VB.Image cqu 
         Height          =   225
         Index           =   0
         Left            =   240
         Picture         =   "測試表單3.frx":5509
         Top             =   1020
         Width           =   300
      End
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Index           =   0
      Left            =   5160
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   720
      Width           =   1935
   End
   Begin MSScriptControlCtl.ScriptControl sc1 
      Left            =   7560
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   6840
      TabIndex        =   5
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Index           =   0
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "測試表單3.frx":5576
      Top             =   5640
      Width           =   3495
   End
   Begin VB.Timer t2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6240
      Top             =   1800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Top             =   4920
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   6960
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpu 
      Height          =   2475
      Left            =   1320
      TabIndex        =   2
      Top             =   3000
      Width           =   3360
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "mini"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   5927
      _cy             =   4366
   End
   Begin UnlightVBE.大人物形像 大人物形像1 
      Height          =   2535
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   4471
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Label1"
      Height          =   615
      Left            =   5040
      TabIndex        =   8
      Top             =   4320
      Width           =   1215
   End
End
Attribute VB_Name = "測試表單3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n() As Integer, n2 As Double, n3 As Integer
Dim ans As String
Dim qnum As Integer
Private Declare Function EbExecuteLine Lib "vba6.dll" ( _
ByVal pStringToExec As Long, _
ByVal Unknownn1 As Long, _
ByVal Unknownn2 As Long, _
ByVal fCheckOnly As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Command1_Click()
'Randomize
'bloodtot = Int(Rnd() * 3) + 0
'Print bloodtot
'main
'大人物形像1.大人物圖片 = App.Path & "\gif\Grunwaldminidown2.png"
'============
'Call 排列組合測試
'============
'Dim m As String
'm = "123456789"
'Mid(m, 3, 1) = "p"
'MsgBox m
'MsgBox (智慧型AI系統類.階層數(6) / 智慧型AI系統類.階層數(6 - 5)) / 智慧型AI系統類.階層數(5)
'MsgBox 智慧型AI系統類.階層數_取C(8, 5)
'智慧型AI系統類.試驗3
'MsgBox Format(Now, "yyyy/m/d_hh:mm:ss")
'=============================================================
'Dim text2 As Object
'Set text2 = Controls.Add("vb.textbox", "text2")
'text2.Move 6360, 360, 2655, 270
'text2.Visible = True
''MsgBox mq.name
''=============
'For i = 1 To 10
'   Load Text2(i)
'   Set Text2(i).Container = 測試表單3
'   Text2(i).Left = 6360
'   Text2(i).Top = 360 + i * 600
'   Text2(i).Text = i
'   Text2(i).Visible = True
'Next
't2.Enabled = True
'測試1.wait 10
'MsgBox n2
'n2 = "123"
'MsgBox Mid(n2, 2, Len(n2))
'=====================================
'For i = 1 To 2
'
'''    '==============
''    Load cge(i)
''    cge(i).Left = cge(i - 1).Left - 100
''    cge(i).Visible = True
'    Load card(i)
'    card(i).Left = card(i - 1).Left + 1000
'    card(i).Visible = True
'    '=============
'    Load cge(i)
'    Set cge(i).Container = card(i)
'    cge(i).Visible = True
'Next
'=====================================
'Dim wwww(1 To 2, 1 To 3, 1 To 4) As Integer
'MsgBox UBound(wwww, 3)
uc擲骰介面1.ZOrder
End Sub
Sub 排列組合測試()
'========假設共qnum張牌
qnum = 10
'===========
ReDim n(1 To Val(qnum))
Dim s As Integer
Dim nm() As String
For i = 1 To qnum   '重設區塊數值
    n(i) = 0
Next
s = 1
'================
Do
    For i = qnum To 1 Step -1
        ans = ans & n(i)
    Next
    '================
    n(1) = n(1) + 1
    排列組合測試_區塊進位 qnum '共[qnum]位數
    '================
    s = s + 1
    ans = ans & "="
Loop Until s > (2 ^ qnum)
nm = Split(ans, "=")
Dim h As Integer
h = 1
For i = 0 To (2 ^ qnum) - 1
    Print nm(i)
    h = h + 1
'    If h > 50 Then
'        Cls
'        h = 1
'    End If
Next

End Sub
Sub 排列組合測試_區塊進位(ByVal num As Integer)
For i = 1 To num - 1
    If n(i) = 2 Then
        n(i + 1) = n(i + 1) + 1
        n(i) = 0
    End If
Next

End Sub

Private Sub Command2_Click()
'Me.Cls
'wmpu.URL = app_path & "mp3\bgm004.mp3"
'wmpu.settings.playCount = 99999
't2.Enabled = True
'sc1.Run "Hello", "andy"
'sc1.Run "test2"
'sc1.Run "test3"
'MsgBox "456"
For i = 1 To Text2.UBound
    Unload Text2(i)
Next
End Sub

Private Sub Command3_Click()
'Dim textlinea As String, str As String, str2(1 To 5, 1 To 5, 1 To 5) As Variant
'Open App.Path & "\test\input1.txt" For Input As #1
'
'Do Until EOF(1)
'   Line Input #1, textlinea
''   stepline textlinea
'   str = str & textlinea & vbCrLf
'Loop
'Close
'sc1.AddCode str
'str2(5, 5, 5) = 1
''MsgBox sc1.Run("test", 2)
'MsgBox sc1.Run("test", str2)
'sc1.reset
'MsgBox sc1.Run("test", str2)
'
'Open App.Path & "\test\input2.txt" For Input As #1
'
'Do Until EOF(1)
'   Line Input #1, textlinea
'   stepline textlinea
'Loop
'Close
'aub = 20
'stepline Text1(0)
'Dim sProcedure As String
'sProcedure = "Sub Hello(sName)" & vbCrLf & _
'                 "    Msgbox ""Hello "" & sName " & _
'                 "End Sub"
'sc1.AddCode sProcedure
Image1.Picture = ImageList1.ListImages(1).Picture
End Sub



Private Sub Form_Click()
MsgBox wmpu.Status
End Sub


Private Sub t2_Timer()
'Print wmpu.Status
'If Left(wmpu.Status, 2) = "播放" Then
'    t2.Enabled = False
'End If
Print "456"
n3 = n3 + 1
If n3 > 10 Then t2.Enabled = False
End Sub
Function stepline(ByVal cmd As String) As Long 'cmd就是vb6代碼
    Dim l As Long '臨時變量,意義不大
    l = EbExecuteLine(StrPtr(ByVal cmd), 0, 0, 0) '這就是實質,簡單吧
    Debug.Print CStr(l) + ":" + cmd '調試用的,無意義
End Function
