'64bit版
Public Declare PtrSafe Function OpenClipboard Lib "user32" (Optional ByVal hwnd As Long = 0) As Long
Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
'32bit版
'Public Declare Function OpenClipboard Lib "user32" (Optional ByVal hwnd As Long = 0) As Long
'Public Declare Function CloseClipboard Lib "user32" () As Long
'Public Declare Function EmptyClipboard Lib "user32" () As Long
'Public Declare  Sub Sleep Lib "kernel32" (ByVal ms As Long)

Public Const VK_SNAPSHOT = &H2C    ''[PrintScrn]キー
Public Const KEYEVENTF_EXTENDEDKEY = &H1    ''キーを押す
Public Const KEYEVENTF_KEYUP = &H2          ''キーを放す

Public Const DST_SHT = "Sheet1"
Public Const START_ROW = 5
Public Const PIC_OFFSET = 3

Public lastPic As Integer



Sub AutoCapture()
    Dim OffsetY As Integer
    Dim CB As Variant
    Dim firstFlg As Boolean
    
    
    destRow = ThisWorkbook.Sheets(DST_SHT).UsedRange.Row
    firstFlg = True
    'クリップボードを空にする。
    OpenClipboard
    EmptyClipboard
    CloseClipboard
    
    MsgBox "AutoCaptureを開始します。" & vbNewLine & _
        "終了するには任意のシートのA1セルにExitと入力してください。", vbInformation
    OffsetY = 1
    Do While True
        CB = Application.ClipboardFormats
        
        If StrConv(ThisWorkbook.Sheets(DST_SHT).Cells(1, 1).Value, vbUpperCase) = "EXIT" Then GoTo Quit
        If CB(1) <> -1 Then
            For i = 1 To UBound(CB)
                If CB(i) = xlClipboardFormatBitmap Then
                    ThisWorkbook.Sheets(DST_SHT).Activate
                    Sleep 500
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
    Loop
    
Quit:
    MsgBox "AutoCaptureを停止しました。", vbInformation
    Sheets(DST_SHT).Cells(1, 1).ClearContents
End Sub


Function getLastRow()
   lastRow = START_ROW
   destRow = ThisWorkbook.Sheets(DST_SHT).UsedRange.Row
   With ThisWorkbook.Sheets(DST_SHT).Shapes
    
    hasPicFlag = False
    For i = 1 To .Count
        Top = .Range(i).Top
        Height = .Range(i).Height
        sharpLastRow = Top + Height
        sharpLastRow = CInt((Top + Height) / .Item(i).TopLeftCell.RowHeight)
        If lastRow <= sharpLastRow Then
            lastPic = i
            lastRow = sharpLastRow
            hasPicFlag = True
        End If
    Next i
    
    If hasPicFlag = True Then
        getLastRow = lastRow + PIC_OFFSET
    Else
        getLastRow = lastRow
    End If
'    Debug.Print getLastRow
    End With
End Function

' 最後の画像の頭の行番号を取得
Function getLastRowHead(destRow)
    If lastPic <> 0 Then
        sharpLastRowHead = destRow - 1
    Else
        sharpLastRowHead = START_ROW
    End If
    getLastRowHead = sharpLastRowHead
End Function



