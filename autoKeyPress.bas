Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, _
ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const SYNCHRONIZE = &H100000
Private Const INFINITE = &HFFFFFFFF
 
Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const KEYEVENTF_KEYUP = &H2
Private Const VK_SNAPSHOT = &H2C
Private Const VK_MENU = &H12

Sub ShellAndWait(pathFile As String)
    With CreateObject("WScript.Shell")
        .Run pathFile, 1, True
    End With
End Sub

 Sub CommandButton1_Click()
    Dim OffsetY As Integer
    Dim CB As Variant
    Dim firstFlg As Boolean
    
 
    
    folderpath = "C:\Windows"
     '通常サイズ
    lngPId = Shell("C:\Windows\Explorer.exe " & folderpath, vbNormalFocus)
    lngPHandle = OpenProcess(SYNCHRONIZE, 0, lngPId)
    If lngPHandle <> 0 Then
        Call WaitForSingleObject(lngPHandle, INFINITE) 'wait for end
        Call CloseHandle(lngPHandle)
    End If
    DoEvents
    Sleep 3
    
    destRow = ThisWorkbook.Sheets(DST_SHT).UsedRange.Row
    firstFlg = True
    
    OpenClipboard
    EmptyClipboard
    CloseClipboard
    keybd_event VK_MENU, 0, 0, 0
    keybd_event VK_SNAPSHOT, 0, 0, 0
    keybd_event VK_SNAPSHOT, 0, KEYEVENTF_KEYUP, 0
    keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0
    
    Debug.Print ("key press")
    ThisWorkbook.Activate
    ThisWorkbook.Sheets(DST_SHT).Activate
    Call startPaste
End Sub



Sub startPaste()
    CB = Application.ClipboardFormats
    If StrConv(ThisWorkbook.Sheets(DST_SHT).Cells(1, 1).Value, vbUpperCase) = "EXIT" Then GoTo Quit
    If CB(1) <> -1 Then
        For i = 1 To UBound(CB)
            If CB(i) = xlClipboardFormatBitmap Then
                Sleep 500
                Debug.Print ("paste start")
                '画像サイズにより位置移動
                destRow = getLastRow()
                '画像最初の場所を見つける
                destRowHead = getLastRowHead(destRow)
                ThisWorkbook.Sheets(DST_SHT).Paste Destination:=ThisWorkbook.Sheets(DST_SHT).Range("B" & destRow)
                ThisWorkbook.Sheets(DST_SHT).Range("A" & destRowHead) = lastPic + 1
                'Debug.Print ("destRowHead = " & destRowHead & " lastPic = " & lastPic + 1)
                'クリップボードを空にする。
                OpenClipboard
                EmptyClipboard
                CloseClipboard
            End If
        Next i
    End If
    DoEvents
    
Quit:
'    MsgBox "AutoCaptureを停止しました。", vbInformation
    Sheets(DST_SHT).Cells(1, 1).ClearContents
End Sub


