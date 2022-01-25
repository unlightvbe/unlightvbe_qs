VERSION 5.00
Begin VB.UserControl uc訊息視窗 
   Appearance      =   0  '平面
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   ClientHeight    =   1365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4920
   BeginProperty Font 
      Name            =   "微軟正黑體"
      Size            =   8.25
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   ScaleHeight     =   1365
   ScaleWidth      =   4920
   Windowless      =   -1  'True
End
Attribute VB_Name = "uc訊息視窗"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim m_message() As String
Private Sub UserControl_Initialize()
ReDim m_message(0) As String
End Sub
Public Property Get MeaageText() As String
   MeaageText = ""
End Property
Public Property Let MeaageText(ByVal New_MeaageText As String)
   Dim i As Integer
   ReDim Preserve m_message(UBound(m_message) + 1) As String
   m_message(UBound(m_message)) = New_MeaageText
   PropertyChanged "MeaageText"
   '=================
   Cls
   If UBound(m_message) <= 5 Then
       For i = 1 To UBound(m_message)
            Print " " & m_message(i)
       Next
   ElseIf UBound(m_message) > 5 Then
       For i = (UBound(m_message) - 5) + 1 To UBound(m_message)
            Print " " & m_message(i)
       Next
   End If
End Property
Sub MessageTextClear()
Cls
ReDim m_message(0) As String
End Sub
