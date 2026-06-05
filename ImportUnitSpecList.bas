Option Explicit

Public Sub ImportUnitSpecList()

    Const TARGET_SHEET  As String = "ユニット仕様一覧"
    Const SRC_SHEET     As String = "ユニット仕様"
    ' ファイル名はプレフィックスで前方一致マッチ
    Const SRC_FILENAME  As String = "3.2_ソフトウェアユニット仕様一覧"

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

    ' ---- ヘッダー行（ソース「ユニット仕様」シートの列名に合わせる）----
    ' A列: 参照ファイルの可変部分(XXXX)、B列以降: ソースの10列
    outputRow = 1
    With wsTarget
        .Cells(outputRow, 1).Value = "参照元識別子"
        .Cells(outputRow, 2).Value = "ソフトウェアコンポーネント仕様ID"
        .Cells(outputRow, 3).Value = "更新No"
        .Cells(outputRow, 4).Value = "変更内容"
        .Cells(outputRow, 5).Value = "仕様"
        .Cells(outputRow, 6).Value = "ソフトウェアユニット仕様ID"
        .Cells(outputRow, 7).Value = "ソフトウェアユニット仕様"
        .Cells(outputRow, 8).Value = "優先順位"
        .Cells(outputRow, 9).Value = "備考"
        .Cells(outputRow, 10).Value = "搭載ソフトウェアユニットID"
        .Cells(outputRow, 11).Value = "搭載ソフトウェアユニット名称"
    End With
    outputRow = 2

    ' ---- フォルダ再帰処理（全データをCollectionに収集）----
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim allRows As Collection
    Set allRows = New Collection
    Call ProcessFolder(fso, fso.GetFolder(unitListFolder), _
                       SRC_FILENAME, SRC_SHEET, allRows)

    ' ---- 一括書き込み ----
    If allRows.Count > 0 Then
        Dim dataArr() As Variant
        ReDim dataArr(1 To allRows.Count, 1 To 11)
        Dim r As Long
        For r = 1 To allRows.Count
            Dim rowArr As Variant
            rowArr = allRows(r)
            Dim c As Long
            For c = 1 To 11
                dataArr(r, c) = rowArr(c)
            Next c
        Next r
        wsTarget.Range(wsTarget.Cells(2, 1), _
                       wsTarget.Cells(1 + allRows.Count, 11)).Value = dataArr
        outputRow = 2 + allRows.Count
    End If

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    ' ---- 詳細シートの書式設定 ----
    Call FormatDetailSheet(wsTarget, outputRow - 1)

    Call BuildSummarySheet(wsTarget)

    MsgBox "インポート完了。" & (outputRow - 2) & " 件の関数仕様を書き込みました。", vbInformation

End Sub

Private Sub ProcessFolder(fso As Object, folder As Object, _
                          srcFileName As String, srcSheetName As String, _
                          allRows As Collection)

    Dim subFolder As Object
    Dim file      As Object
    Dim normalizedTarget As String
    Dim ignoreFolders As Variant
    Dim ignoreName As String

    normalizedTarget = NormalizeFileName(srcFileName)
    ignoreFolders = Array(".svn")

    For Each file In folder.Files
        If Left(file.Name, 1) <> "~" Then
            ' プレフィックス前方一致でマッチ
            If Left(NormalizeFileName(file.Name), Len(normalizedTarget)) = normalizedTarget Then
                Call ReadMainFile(file.Path, srcSheetName, allRows)
            End If
        End If
    Next file

    For Each subFolder In folder.SubFolders
        ignoreName = NormalizeFolderName(subFolder.Name)
        If Not IsIgnoredFolder(ignoreName, ignoreFolders) Then
            ' アクセス不可フォルダ（シンボリックリンク等）をスキップ
            If fso.FolderExists(subFolder.Path) Then
                On Error Resume Next
                Call ProcessFolder(fso, subFolder, srcFileName, srcSheetName, allRows)
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
                          allRows As Collection)

    Dim wbSrc   As Workbook
    Dim wsSrc   As Worksheet
    Dim ws      As Worksheet
    Dim lastRow As Long
    Dim i       As Long
    Dim col     As Long

    ' ファイル名から可変部分(XXXX)を抽出する
    ' 例: "3.2_ソフトウェアユニット仕様一覧_XXXX.xlsx" → "XXXX"
    Const FILE_PREFIX As String = "3.2_ソフトウェアユニット仕様一覧_"
    Dim fileName    As String
    Dim fileIdent   As String
    fileName  = Mid$(filePath, InStrRev(filePath, "\") + 1)
    ' 拡張子を除去
    If InStrRev(fileName, ".") > 0 Then
        fileName = Left$(fileName, InStrRev(fileName, ".") - 1)
    End If
    ' プレフィックスより後ろを取得
    If Left$(fileName, Len(FILE_PREFIX)) = FILE_PREFIX Then
        fileIdent = Mid$(fileName, Len(FILE_PREFIX) + 1)
    Else
        fileIdent = fileName
    End If

    Set wbSrc = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)

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

    ' 行1 = ヘッダー、行2以降がデータ（フラット構造）
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row

    ' ソース全データを一括で配列に読み込む
    Dim srcArr As Variant
    If lastRow >= 2 Then
        srcArr = wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(lastRow, 10)).Value
    End If

    wbSrc.Close SaveChanges:=False

    If IsEmpty(srcArr) Then Exit Sub
    If Not IsArray(srcArr) Then Exit Sub

    For i = 1 To UBound(srcArr, 1)
        ' ソフトウェアユニット仕様ID(E列=srcArr列5)が空の行はスキップ
        If Trim(CStr(srcArr(i, 5))) <> "" Then
            Dim rowArr(1 To 11) As Variant
            rowArr(1) = fileIdent
            For col = 1 To 10
                rowArr(col + 1) = Trim(CStr(srcArr(i, col)))
            Next col
            allRows.Add rowArr
        End If
    Next i

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
        .Cells(outRow, 3).Value = "仕様"
        .Cells(outRow, 4).Value = "コンポーネント仕様ID一覧"
        .Cells(outRow, 5).Value = "搭載ソフトウェアユニットID一覧"
        .Cells(outRow, 6).Value = "搭載ソフトウェアユニット名称一覧"
    End With
    outRow = 2

    ' ---- 詳細シートを走査（行2以降がデータ、行1はヘッダー）----
    ' 詳細シートはcol1=参照元識別子, col6=ソフトウェアユニット仕様ID(+1シフト)
    lastRow = wsDetail.Cells(wsDetail.Rows.Count, 6).End(xlUp).Row

    i = 2
    Do While i <= lastRow

        Dim curUnitID   As String
        Dim curUnitName As String
        Dim curStatus   As String
        Dim funcIDs     As String
        Dim funcNames   As String

        ' 集約キー: col6 = ソフトウェアユニット仕様ID
        curUnitID   = Trim(CStr(wsDetail.Cells(i, 6).Value))
        curUnitName = Trim(CStr(wsDetail.Cells(i, 7).Value))
        curStatus   = Trim(CStr(wsDetail.Cells(i, 5).Value))
        funcIDs     = ""
        funcNames   = ""

        Dim funcCarriedNames As String
        funcCarriedNames = ""

        ' 同じ ユニット仕様ID(col6) が続く間まとめる
        Do While i <= lastRow And _
                 Trim(CStr(wsDetail.Cells(i, 6).Value)) = curUnitID

            Dim fID   As String
            Dim fName As String
            Dim fCarriedName As String
            ' col2=コンポーネント仕様ID, col10=搭載ユニットID, col11=搭載ユニット名称
            fID          = Trim(CStr(wsDetail.Cells(i, 2).Value))
            fName        = Trim(CStr(wsDetail.Cells(i, 10).Value))
            fCarriedName = Trim(CStr(wsDetail.Cells(i, 11).Value))

            If fID <> "" Then
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

        ' サマリー行に書き出し（件数列なし）
        With wsSummary
            .Cells(outRow, 1).Value = curUnitID
            .Cells(outRow, 2).Value = curUnitName
            .Cells(outRow, 3).Value = curStatus
            .Cells(outRow, 4).Value = funcIDs
            .Cells(outRow, 5).Value = funcNames
            .Cells(outRow, 6).Value = funcCarriedNames

            .Cells(outRow, 4).WrapText = True
            .Cells(outRow, 5).WrapText = True
            .Cells(outRow, 6).WrapText = True

            ' 行高さを内容に合わせる
            .Rows(outRow).AutoFit
        End With

        outRow = outRow + 1

    Loop

    ' ---- 列幅を整える ----
    wsSummary.Columns("A:F").AutoFit

    ' 関数一覧列は広げすぎず上限を設ける
    Dim c As Integer
    For c = 4 To 6
        If wsSummary.Columns(c).ColumnWidth > 50 Then
            wsSummary.Columns(c).ColumnWidth = 50
        End If
    Next c

    ' ---- 1行目ヘッダーの書式（データのある列のみ）----
    Dim lastHeaderCol As Long
    lastHeaderCol = wsSummary.Cells(1, wsSummary.Columns.Count).End(xlToLeft).Column
    With wsSummary.Range(wsSummary.Cells(1, 1), wsSummary.Cells(1, lastHeaderCol))
        .Interior.Color = RGB(0, 112, 192)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    ' ---- 罫線（データが存在する範囲全体）----
    If outRow > 2 Then
        With wsSummary.Range(wsSummary.Cells(1, 1), wsSummary.Cells(outRow - 1, lastHeaderCol)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
    End If

End Sub

' ----------------------------------------------------------
' 「ユニット仕様一覧」シートの書式を設定する
' ----------------------------------------------------------
Private Sub FormatDetailSheet(ws As Worksheet, ByVal dataLastRow As Long)

    ' ---- 1行目ヘッダーの書式（データのある列のみ）----
    Dim lastHeaderCol As Long
    lastHeaderCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastHeaderCol))
        .Interior.Color = RGB(0, 112, 192)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    ' ---- 列幅・折り返し ----
    ws.Columns("B").ColumnWidth = 65
    ws.Columns("B").WrapText = True

    ws.Columns("D").ColumnWidth = 50

    ws.Columns("F").ColumnWidth = 120
    ws.Columns("F").WrapText = True

    ' ---- 罫線（データが存在する範囲全体）----
    If dataLastRow >= 1 Then
        With ws.Range(ws.Cells(1, 1), ws.Cells(dataLastRow, 11)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
    End If

End Sub
