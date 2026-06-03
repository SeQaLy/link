Option Explicit

Public Sub ImportUnitSpecList()

    Const TARGET_SHEET  As String = "ユニット仕様一覧"
    Const SRC_SHEET     As String = "3.2.ユニット仕様一覧"
    Const SRC_FILENAME  As String = "XXXXXX"

    Dim fso             As Object
    Dim unitListFolder  As String
    Dim wsTarget        As Worksheet
    Dim outputRow       As Long
    Dim ws              As Worksheet

    unitListFolder = ThisWorkbook.Path & "\UnitList"

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(unitListFolder) Then
        MsgBox "UnitListフォルダが見つかりません:" & vbCrLf & unitListFolder, vbCritical
        Exit Sub
    End If

    ' ---- ターゲットシートの準備 ----
    Set wsTarget = Nothing
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = TARGET_SHEET Then
            Set wsTarget = ws
            Exit For
        End If
    Next ws

    If wsTarget Is Nothing Then
        Set wsTarget = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsTarget.Name = TARGET_SHEET
    Else
        wsTarget.Cells.Clear
    End If

    ' ---- ヘッダー行 ----
    outputRow = 1
    With wsTarget
        .Cells(outputRow, 1).Value = "ソフトウェアユニット仕様ID"
        .Cells(outputRow, 2).Value = "ソフトウェアユニット仕様名"
        .Cells(outputRow, 3).Value = "仕様status"
        .Cells(outputRow, 4).Value = "関数仕様ID"
        .Cells(outputRow, 5).Value = "関数仕様名"
        .Cells(outputRow, 6).Value = "優先順位"
        .Cells(outputRow, 7).Value = "備考"
        .Cells(outputRow, 8).Value = "搭載関数ID"
        .Cells(outputRow, 9).Value = "搭載関数名"
    End With
    outputRow = 2

    ' ---- フォルダ再帰処理 ----
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Call ProcessFolder(fso, fso.GetFolder(unitListFolder), _
                       SRC_FILENAME, SRC_SHEET, wsTarget, outputRow)

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    Call BuildSummarySheet(wsTarget)

    MsgBox "インポート完了。" & (outputRow - 2) & " 件の関数仕様を書き込みました。", vbInformation

End Sub

Private Sub ProcessFolder(fso As Object, folder As Object, _
                          srcFileName As String, srcSheetName As String, _
                          wsTarget As Worksheet, ByRef outputRow As Long)

    Dim subFolder As Object
    Dim file      As Object
    Dim normalizedTarget As String
    Dim ignoreFolders As Variant
    Dim ignoreName As String

    normalizedTarget = NormalizeFileName(srcFileName)
    ignoreFolders = Array(".svn")

    For Each file In folder.Files
        If Left(file.Name, 1) <> "~" Then
            If NormalizeFileName(file.Name) = normalizedTarget Then
                Call ReadMainFile(file.Path, srcSheetName, wsTarget, outputRow)
            End If
        End If
    Next file

    For Each subFolder In folder.SubFolders
        ignoreName = NormalizeFolderName(subFolder.Name)
        If Not IsIgnoredFolder(ignoreName, ignoreFolders) Then
            ' アクセス不可フォルダ（シンボリックリンク等）をスキップ
            If fso.FolderExists(subFolder.Path) Then
                On Error Resume Next
                Call ProcessFolder(fso, subFolder, srcFileName, srcSheetName, wsTarget, outputRow)
                On Error GoTo 0
            End If
        End If
    Next subFolder

End Sub

Private Function NormalizeFolderName(ByVal folderName As String) As String

    NormalizeFolderName = LCase$(Trim$(folderName))

End Function

Private Function IsIgnoredFolder(ByVal folderName As String, ByVal ignoreFolders As Variant) As Boolean

    Dim i As Long

    For i = LBound(ignoreFolders) To UBound(ignoreFolders)
        If folderName = NormalizeFolderName(CStr(ignoreFolders(i))) Then
            IsIgnoredFolder = True
            Exit Function
        End If
    Next i

End Function

Private Function NormalizeFileName(ByVal fileName As String) As String

    Dim result As String
    result = LCase$(Trim$(fileName))

    ' 拡張子を外して比較する
    If InStrRev(result, ".") > 0 Then
        result = Left$(result, InStrRev(result, ".") - 1)
    End If

    ' 比較に不要な空白類を除去する
    result = Replace(result, " ", "")
    result = Replace(result, ChrW(&H3000), "")
    result = Replace(result, vbTab, "")
    result = Replace(result, vbCr, "")
    result = Replace(result, vbLf, "")

    NormalizeFileName = result

End Function

Private Sub ReadMainFile(filePath As String, srcSheetName As String, _
                          wsTarget As Worksheet, ByRef outputRow As Long)

    Dim wbSrc       As Workbook
    Dim wsSrc       As Worksheet
    Dim ws          As Worksheet
    Dim lastRow     As Long
    Dim i           As Long

    Dim unitID      As String
    Dim unitName    As String
    Dim unitStatus  As String

    Set wbSrc = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)

    ' 対象シートを探す
    Set wsSrc = Nothing
    For Each ws In wbSrc.Worksheets
        If ws.Name = srcSheetName Then
            Set wsSrc = ws
            Exit For
        End If
    Next ws

    If wsSrc Is Nothing Then
        wbSrc.Close SaveChanges:=False
        Exit Sub
    End If

    ' 最終行をA列・C列のどちらか大きい方で判定
    Dim lastA As Long, lastC As Long
    lastA = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    lastC = wsSrc.Cells(wsSrc.Rows.Count, 3).End(xlUp).Row
    lastRow = IIf(lastA > lastC, lastA, lastC)

    unitID     = ""
    unitName   = ""
    unitStatus = ""

    ' 行4以降がデータ行（行1〜3はヘッダー）
    For i = 4 To lastRow

        Dim cellA As String
        Dim cellC As String
        cellA = Trim(CStr(wsSrc.Cells(i, 1).Value))
        cellC = Trim(CStr(wsSrc.Cells(i, 3).Value))

        If cellA <> "" Then
            ' ユニット仕様ヘッダー行
            unitID     = cellA
            unitName   = Trim(CStr(wsSrc.Cells(i, 4).Value))
            unitStatus = Trim(CStr(wsSrc.Cells(i, 2).Value))

        ElseIf cellC <> "" Then
            ' 関数仕様行 → ターゲットへ1行書き出し
            With wsTarget
                .Cells(outputRow, 1).Value = unitID
                .Cells(outputRow, 2).Value = unitName
                .Cells(outputRow, 3).Value = unitStatus
                .Cells(outputRow, 4).Value = cellC
                .Cells(outputRow, 5).Value = Trim(CStr(wsSrc.Cells(i, 4).Value))
                .Cells(outputRow, 6).Value = Trim(CStr(wsSrc.Cells(i, 5).Value))
                .Cells(outputRow, 7).Value = Trim(CStr(wsSrc.Cells(i, 6).Value))
                .Cells(outputRow, 8).Value = Trim(CStr(wsSrc.Cells(i, 7).Value))
                .Cells(outputRow, 9).Value = Trim(CStr(wsSrc.Cells(i, 8).Value))
            End With
            outputRow = outputRow + 1
        End If

    Next i

    wbSrc.Close SaveChanges:=False

End Sub


Private Sub BuildSummarySheet(wsDetail As Worksheet)

    Const SUMMARY_SHEET As String = "ユニット仕様サマリー"

    Dim wsSummary   As Worksheet
    Dim ws          As Worksheet
    Dim lastRow     As Long
    Dim i           As Long
    Dim outRow      As Long

    ' ---- サマリーシートの準備 ----
    Set wsSummary = Nothing
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SUMMARY_SHEET Then
            Set wsSummary = ws
            Exit For
        End If
    Next ws

    If wsSummary Is Nothing Then
        Set wsSummary = ThisWorkbook.Worksheets.Add( _
            After:=wsDetail)
        wsSummary.Name = SUMMARY_SHEET
    Else
        wsSummary.Cells.Clear
    End If

    ' ---- ヘッダー ----
    outRow = 1
    With wsSummary
        .Cells(outRow, 1).Value = "ソフトウェアユニット仕様ID"
        .Cells(outRow, 2).Value = "ソフトウェアユニット仕様名"
        .Cells(outRow, 3).Value = "仕様status"
        .Cells(outRow, 4).Value = "関数数"
        .Cells(outRow, 5).Value = "関数仕様ID一覧"
        .Cells(outRow, 6).Value = "関数仕様名一覧"
        .Cells(outRow, 7).Value = "搭載関数名一覧"
    End With
    outRow = 2

    ' ---- 詳細シートを走査（行2以降がデータ、行1はヘッダー）----
    lastRow = wsDetail.Cells(wsDetail.Rows.Count, 1).End(xlUp).Row

    i = 2
    Do While i <= lastRow

        Dim curUnitID   As String
        Dim curUnitName As String
        Dim curStatus   As String
        Dim funcCount   As Long
        Dim funcIDs     As String
        Dim funcNames   As String

        curUnitID   = Trim(CStr(wsDetail.Cells(i, 1).Value))
        curUnitName = Trim(CStr(wsDetail.Cells(i, 2).Value))
        curStatus   = Trim(CStr(wsDetail.Cells(i, 3).Value))
        funcCount   = 0
        funcIDs     = ""
        funcNames   = ""

        Dim funcCarriedNames As String
        funcCarriedNames = ""

        ' 同じ ユニット仕様ID が続く間まとめる
        Do While i <= lastRow And _
                 Trim(CStr(wsDetail.Cells(i, 1).Value)) = curUnitID

            Dim fID   As String
            Dim fName As String
            Dim fCarriedName As String
            fID          = Trim(CStr(wsDetail.Cells(i, 4).Value))
            fName        = Trim(CStr(wsDetail.Cells(i, 5).Value))
            fCarriedName = Trim(CStr(wsDetail.Cells(i, 9).Value))

            If fID <> "" Then
                funcCount = funcCount + 1
                If funcIDs = "" Then
                    funcIDs          = fID
                    funcNames        = fName
                    funcCarriedNames = fCarriedName
                Else
                    funcIDs          = funcIDs          & Chr(10) & fID
                    funcNames        = funcNames        & Chr(10) & fName
                    funcCarriedNames = funcCarriedNames & Chr(10) & fCarriedName
                End If
            End If

            i = i + 1
        Loop

        ' サマリー行に書き出し
        With wsSummary
            .Cells(outRow, 1).Value = curUnitID
            .Cells(outRow, 2).Value = curUnitName
            .Cells(outRow, 3).Value = curStatus
            .Cells(outRow, 4).Value = funcCount
            .Cells(outRow, 5).Value = funcIDs
            .Cells(outRow, 6).Value = funcNames
            .Cells(outRow, 7).Value = funcCarriedNames

            ' 折り返し表示（改行が見えるように）
            .Cells(outRow, 5).WrapText = True
            .Cells(outRow, 6).WrapText = True
            .Cells(outRow, 7).WrapText = True

            ' 行高さを内容に合わせる
            .Rows(outRow).AutoFit
        End With

        outRow = outRow + 1

    Loop

    ' ---- 列幅を整える ----
    wsSummary.Columns("A:F").AutoFit

    ' 関数一覧列は広げすぎず上限を設ける
    Dim c As Integer
    For c = 5 To 7
        If wsSummary.Columns(c).ColumnWidth > 50 Then
            wsSummary.Columns(c).ColumnWidth = 50
        End If
    Next c

End Sub
