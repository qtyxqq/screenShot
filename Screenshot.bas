'64bit版
Private Declare PtrSafe Function OpenClipboard Lib "user32" (Optional ByVal hwnd As Long = 0) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
'32bit版
'Private Declare Function OpenClipboard Lib "user32" (Optional ByVal hwnd As Long = 0) As Long
'Private Declare Function CloseClipboard Lib "user32" () As Long
'Private Declare Function EmptyClipboard Lib "user32" () As Long
'Private Declare  Sub Sleep Lib "kernel32" (ByVal ms As Long)

Const VK_SNAPSHOT = &H2C    ''[PrintScrn]キー
Const KEYEVENTF_EXTENDEDKEY = &H1    ''キーを押す
Const KEYEVENTF_KEYUP = &H2          ''キーを放す

Const DST_SHT = "Sheet1"
Const START_ROW = 5
Const PIC_OFFSET = 3


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
                     ThisWorkbook.Sheets(DST_SHT).Paste Destination:=ThisWorkbook.Sheets(DST_SHT).Range("B" & destRow)

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
            lastRow = sharpLastRow
            hasPicFlag = True
        End If
    Next i
    
    If hasPicFlag = True Then
        getLastRow = lastRow + PIC_OFFSET
    Else
        getLastRow = lastRow
    End If
    Debug.Print getLastRow
    End With
End Function


