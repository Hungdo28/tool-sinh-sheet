Option Explicit

' ======================================================================
' MACRO: SINH SHEET TEST CASE TU DONG
' ======================================================================

Public Sub Tao_Sheet_Testcase()
    Dim wsConfig As Worksheet, wsSource As Worksheet, wsOverview As Worksheet
    Dim lastRow As Long, lastCol As Long, r As Long, c As Long
    Dim headerRow As Long, groupCol As Long
    Dim groupColInput As Variant
    Dim forwardFill As Boolean, createOverview As Boolean, addBackLink As Boolean
    Dim rawVal As Variant, currentGroup As String, lastGroup As String
    Dim cleanName As String
    Dim dictGroups As Object
    Dim dictCounts As Object
    Dim key As Variant
    Dim targetWs As Worksheet
    Dim targetRow As Long
    Dim totalGroups As Long, totalTC As Long
    Dim i As Long
    Dim screenUpdateState As Boolean, alertState As Boolean

    On Error GoTo ErrorHandler

    Set wsConfig = ThisWorkbook.Sheets("HUONG_DAN_&_CAU_HINH")
    Set wsSource = ThisWorkbook.Sheets("Testcase_Tong")

    ' Doc thong so tu sheet huong dan
    headerRow = CLng(wsConfig.Range("C7").Value)
    If headerRow < 1 Then headerRow = 1

    groupColInput = wsConfig.Range("C8").Value
    If IsNumeric(groupColInput) Then
        groupCol = CLng(groupColInput)
    Else
        groupCol = wsSource.Cells(headerRow, CStr(groupColInput)).Column
    End If
    If groupCol < 1 Then groupCol = 2

    forwardFill = (UCase(Trim(wsConfig.Range("C9").Value)) = "CO" Or UCase(Trim(wsConfig.Range("C9").Value)) = "CÓ" Or UCase(Trim(wsConfig.Range("C9").Value)) = "YES")
    createOverview = (UCase(Trim(wsConfig.Range("C10").Value)) = "CO" Or UCase(Trim(wsConfig.Range("C10").Value)) = "CÓ" Or UCase(Trim(wsConfig.Range("C10").Value)) = "YES")
    addBackLink = (UCase(Trim(wsConfig.Range("C11").Value)) = "CO" Or UCase(Trim(wsConfig.Range("C11").Value)) = "CÓ" Or UCase(Trim(wsConfig.Range("C11").Value)) = "YES")

    screenUpdateState = Application.ScreenUpdating
    alertState = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    lastRow = wsSource.Cells(wsSource.Rows.Count, groupCol).End(xlUp).Row
    If lastRow <= headerRow Then
        lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    End If

    If lastRow <= headerRow Then
        Application.ScreenUpdating = screenUpdateState
        Application.DisplayAlerts = alertState
        MsgBox "Khong tim thay du lieu testcase ben duoi dong tieu de tai sheet Testcase_Tong!", vbExclamation, "Thong Bao"
        Exit Sub
    End If

    lastCol = wsSource.Cells(headerRow, wsSource.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = 13

    ' Quet phan nhom testcase
    Set dictGroups = CreateObject("Scripting.Dictionary")
    Set dictCounts = CreateObject("Scripting.Dictionary")
    lastGroup = ""

    For r = headerRow + 1 To lastRow
        If Application.WorksheetFunction.CountA(wsSource.Range(wsSource.Cells(r, 1), wsSource.Cells(r, lastCol))) > 0 Then
            rawVal = wsSource.Cells(r, groupCol).Value

            If wsSource.Cells(r, groupCol).MergeCells Then
                rawVal = wsSource.Cells(r, groupCol).MergeArea.Cells(1, 1).Value
            End If

            If Trim(CStr(rawVal)) <> "" Then
                currentGroup = Trim(CStr(rawVal))
                lastGroup = currentGroup
            ElseIf forwardFill And lastGroup <> "" Then
                currentGroup = lastGroup
            Else
                currentGroup = "Chua_Phan_Nhom"
            End If

            If Not dictGroups.Exists(currentGroup) Then
                dictGroups.Add currentGroup, CStr(r)
                dictCounts.Add currentGroup, 1
            Else
                dictGroups(currentGroup) = dictGroups(currentGroup) & "," & CStr(r)
                dictCounts(currentGroup) = dictCounts(currentGroup) + 1
            End If
        End If
    Next r

    If dictGroups.Count = 0 Then
        Application.ScreenUpdating = screenUpdateState
        Application.DisplayAlerts = alertState
        MsgBox "Khong tim thay nhom testcase nao de sinh sheet!", vbExclamation, "Thong Bao"
        Exit Sub
    End If

    ' Xoa cac sheet con cu
    Dim ws As Worksheet, keepDict As Object
    Set keepDict = CreateObject("Scripting.Dictionary")
    keepDict.Add "HUONG_DAN_&_CAU_HINH", True
    keepDict.Add "Testcase_Tong", True

    For Each ws In ThisWorkbook.Worksheets
        If Not keepDict.Exists(ws.Name) Then
            ws.Delete
        End If
    Next ws

    ' Tao cac sheet con
    totalGroups = dictGroups.Count
    totalTC = 0

    For Each key In dictGroups.Keys
        cleanName = Chuan_Hoa_Ten_Sheet(CStr(key))
        Set targetWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        targetWs.Name = cleanName

        ' Sao chep do rong cot
        For c = 1 To lastCol
            targetWs.Columns(c).ColumnWidth = wsSource.Columns(c).ColumnWidth
        Next c

        targetRow = 1

        ' Them link quay lai Tong quan
        If createOverview And addBackLink Then
            With targetWs.Cells(targetRow, 1)
                .Value = "<- Quay lai Trang Tong quan (Overview)"
                .Font.Name = "Segoe UI"
                .Font.Size = 10
                .Font.Bold = True
                .Font.Color = RGB(29, 78, 216)
                .Font.Underline = xlUnderlineStyleSingle
                targetWs.Hyperlinks.Add Anchor:=targetWs.Cells(targetRow, 1), Address:="", SubAddress:="'Tong_Quan'!A1"
            End With
            targetWs.Rows(targetRow).RowHeight = 24
            targetRow = targetRow + 1
        End If

        ' Sao chep Header
        For r = 1 To headerRow
            wsSource.Rows(r).Copy targetWs.Rows(targetRow)
            targetRow = targetRow + 1
        Next r

        ' Sao chep cac dong testcase con
        Dim rowArr() As String
        rowArr = Split(dictGroups(key), ",")
        For i = LBound(rowArr) To UBound(rowArr)
            Dim srcRow As Long
            srcRow = CLng(rowArr(i))
            wsSource.Rows(srcRow).Copy targetWs.Rows(targetRow)

            If Trim(CStr(targetWs.Cells(targetRow, groupCol).Value)) = "" Then
                targetWs.Cells(targetRow, groupCol).Value = key
            End If

            targetRow = targetRow + 1
            totalTC = totalTC + 1
        Next i
    Next key

    ' Tao trang Tong quan
    If createOverview Then
        Set wsOverview = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets("Testcase_Tong"))
        wsOverview.Name = "Tong_Quan"
        Tao_Bang_Tong_Quan wsOverview, dictCounts, totalTC
    End If

    wsConfig.Activate
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "DA SINH SHEET TEST CASE THANH CONG!" & vbCrLf & vbCrLf & _
           "- So luong Sheet con da tao: " & totalGroups & vbCrLf & _
           "- Tong so Test Case con: " & totalTC & vbCrLf & _
           "- Bang Tong quan (Tong_Quan) da duoc tao day du hyperlink.", _
           vbInformation, "Hoan Tat"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Loi khi sinh sheet: " & Err.Description, vbCritical, "Loi"
End Sub

Private Function Chuan_Hoa_Ten_Sheet(ByVal rawName As String) As String
    Dim s As String, i As Long
    Dim forbidden As String
    forbidden = "\/?*[]:"
    s = Trim(rawName)
    For i = 1 To Len(forbidden)
        s = Replace(s, Mid(forbidden, i, 1), "_")
    Next i
    s = Replace(s, "'", "")
    If Trim(s) = "" Then s = "Sheet"
    If Len(s) > 28 Then s = Left(s, 28)

    Dim candidate As String, counter As Long
    candidate = s
    counter = 1
    Do While Sheet_Ton_Tai(candidate)
        candidate = Left(s, 26) & "_" & counter
        counter = counter + 1
    Loop
    Chuan_Hoa_Ten_Sheet = candidate
End Function

Private Function Sheet_Ton_Tai(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    Sheet_Ton_Tai = Not (ws Is Nothing)
    On Error GoTo 0
End Function

Private Sub Tao_Bang_Tong_Quan(ws As Worksheet, dictCounts As Object, totalTC As Long)
    Dim r As Long, stt As Long
    Dim key As Variant, sName As String, cnt As Long

    ws.Cells(2, 2).Value = "BANG TONG QUAN TEST CASES THEO NHOM / MODULE"
    ws.Cells(2, 2).Font.Name = "Segoe UI"
    ws.Cells(2, 2).Font.Size = 15
    ws.Cells(2, 2).Font.Bold = True
    ws.Cells(2, 2).Font.Color = RGB(30, 58, 138)

    ws.Cells(3, 2).Value = "Thoi gian tao: " & Format(Now, "yyyy-mm-dd hh:mm:ss") & " | Tong so testcase: " & totalTC
    ws.Cells(3, 2).Font.Name = "Segoe UI"
    ws.Cells(3, 2).Font.Size = 10
    ws.Cells(3, 2).Font.Italic = True
    ws.Cells(3, 2).Font.Color = RGB(100, 116, 139)

    ws.Cells(5, 2).Value = "STT"
    ws.Cells(5, 3).Value = "Ten Nhom / Module"
    ws.Cells(5, 4).Value = "Ten Sheet Chi Tiet (Click de mo)"
    ws.Cells(5, 5).Value = "So Test Case Con"
    ws.Cells(5, 6).Value = "Ty Le (%)"

    With ws.Range("B5:F5")
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(30, 58, 138)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 28
    End With

    r = 6
    stt = 1
    For Each key In dictCounts.Keys
        cnt = dictCounts(key)
        sName = Chuan_Hoa_Ten_Sheet(CStr(key))

        ws.Cells(r, 2).Value = stt
        ws.Cells(r, 2).HorizontalAlignment = xlCenter

        ws.Cells(r, 3).Value = key

        ws.Cells(r, 4).Value = "Sheet: " & sName
        ws.Hyperlinks.Add Anchor:=ws.Cells(r, 4), Address:="", SubAddress:="'" & sName & "'!A1"
        ws.Cells(r, 4).Font.Name = "Segoe UI"
        ws.Cells(r, 4).Font.Size = 10
        ws.Cells(r, 4).Font.Color = RGB(37, 99, 235)
        ws.Cells(r, 4).Font.Underline = xlUnderlineStyleSingle

        ws.Cells(r, 5).Value = cnt
        ws.Cells(r, 5).HorizontalAlignment = xlCenter

        If totalTC > 0 Then
            ws.Cells(r, 6).Value = cnt / totalTC
            ws.Cells(r, 6).NumberFormat = "0.0%"
        Else
            ws.Cells(r, 6).Value = "0%"
        End If
        ws.Cells(r, 6).HorizontalAlignment = xlCenter

        If stt Mod 2 = 0 Then
            ws.Range(ws.Cells(r, 2), ws.Cells(r, 6)).Interior.Color = RGB(248, 250, 252)
        End If

        ws.Range(ws.Cells(r, 2), ws.Cells(r, 6)).Borders.LineStyle = 1
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 6)).Borders.Color = RGB(203, 213, 225)
        ws.Rows(r).RowHeight = 22

        r = r + 1
        stt = stt + 1
    Next key

    ' Tong cong
    ws.Cells(r, 2).Value = "TONG CONG"
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 4)).Merge
    ws.Cells(r, 2).HorizontalAlignment = xlCenter
    ws.Cells(r, 5).Value = totalTC
    ws.Cells(r, 5).HorizontalAlignment = xlCenter
    ws.Cells(r, 6).Value = "100.0%"
    ws.Cells(r, 6).HorizontalAlignment = xlCenter

    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 6))
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Interior.Color = RGB(241, 245, 249)
        .Borders.LineStyle = 1
        .Borders.Color = RGB(148, 163, 184)
        .RowHeight = 24
    End With

    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 8
    ws.Columns("C").ColumnWidth = 36
    ws.Columns("D").ColumnWidth = 35
    ws.Columns("E").ColumnWidth = 18
    ws.Columns("F").ColumnWidth = 15
End Sub

Public Sub Xoa_Sheet_Con()
    Dim ws As Worksheet
    Dim ans As VbMsgBoxResult
    ans = MsgBox("Ban co chac chan muon xoa tat ca cac sheet con da sinh ra khong?" & vbCrLf & _
                 "(Sheet 'HUONG_DAN_&_CAU_HINH' va 'Testcase_Tong' se luon duoc giu lai)", vbQuestion + vbYesNo, "Xac Nhan")
    If ans <> vbYes Then Exit Sub

    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "HUONG_DAN_&_CAU_HINH" And ws.Name <> "Testcase_Tong" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
    MsgBox "Da don dep cac sheet con thanh cong!", vbInformation, "Thong Bao"
End Sub

Public Sub Chuyen_Toi_Sheet_Testcase()
    On Error Resume Next
    ThisWorkbook.Sheets("Testcase_Tong").Activate
End Sub
