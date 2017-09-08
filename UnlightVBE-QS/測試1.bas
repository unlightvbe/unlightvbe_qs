Attribute VB_Name = "測試1"
Sub 測試模組程式()
'MsgBox "這是基本測試"
'測試表單2.Label1.Caption = "1111"
End Sub

Sub 測試開始選項()
Dim i, j As Integer
For i = 1 To 3
   Formgamesetting.personnameus(i).ListIndex = 0
   Formgamesetting.personnamecom(i).ListIndex = 0
Next
Formgamesetting.opnpersonvs(2).Value = True
End Sub

Public Sub test2()
MsgBox "test2"
End Sub
Public Sub SearchDirectory()
  Dim mypath As String, mydir As String
  Dim DirectoryBuff()
  Dim Index As Integer
  Index = 0
    mypath = "g:\"
    mydir = Dir(mypath, vbDirectory) ' 找尋第一個子目錄。
    ReDim DirectoryBuff(0)
  Do While True
    Do While mydir <> ""
     ' 跳過目前的目錄及上層目錄。
     If mydir <> "." And mydir <> ".." Then
     ' 使用位元比對來確定 MyName 代表一目錄。
     If (GetAttr(mypath & mydir) And vbDirectory) = vbDirectory Then
     Debug.Print mydir ' 將目錄名稱顯示出來。
     ReDim Preserve DirectoryBuff(UBound(DirectoryBuff) + 1)
     DirectoryBuff(UBound(DirectoryBuff)) = mypath + mydir
     End If
     End If
     mydir = Dir()
    Loop
    Index = Index + 1
    If Index > UBound(DirectoryBuff) Then Exit Do
    mypath = DirectoryBuff(Index) + "\"
    mydir = Dir(mypath, vbDirectory)
  Loop
End Sub
Public Function GetExtName(strFileName As String) As String
    Dim strTmp As String
    Dim strByte As String
    Dim i As Long
    For i = Len(strFileName) To 1 Step -1
        strByte = Mid(strFileName, i, 1)
        If strByte <> "." Then
            strTmp = strByte + strTmp
        Else
            Exit For
        End If
    Next i
    GetExtName = strTmp
End Function
Public Sub wait(Optional ByVal sgnSecondToDelay As Single)
Dim sgnThisTime As Single, sgnCount As Single

      If sgnSecondToDelay = 0 Then
         Exit Sub
      Else
         If sgnSecondToDelay < 0.01 Then
              MsgBox "延遲時間無法小於 0.01 秒", vbOKOnly, "參數錯誤"
            Exit Sub
         End If
      End If

     '主要延遲迴圈
     sgnThisTime = Timer
     Do While sgnCount < sgnSecondToDelay
        sgnCount = Timer - sgnThisTime
        DoEvents
     Loop
End Sub
