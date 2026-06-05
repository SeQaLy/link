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
    Application.Calculation = xlCalculationManual

    Dim allRows As Collection
    Set allRows = New Collection
    Dim ignoreFolders As Variant
    ignoreFolders = Array(".svn")
    ' プレフィックスは拡張子除去不要のため NormalizePrefix を使う
    Call ProcessFolder(fso, fso.GetFolder(unitListFolder), _
                       NormalizePrefix(SRC_FILENAME), SRC_SHEET, ignoreFolders, allRows)

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
    Application.Calculation = xlCalculationAutomatic

    ' ---- 詳細シートの書式設定 ----
    Call FormatDetailSheet(wsTarget, outputRow - 1)

    Call BuildSummarySheet(wsTarget)

    MsgBox "インポート完了。" & (outputRow - 2) & " 件の関数仕様を書き込みました。", vbInformation

End Sub

Private Sub ProcessFolder(fso As Object, folder As Object, _
                          normalizedTarget As String, srcSheetName As String, _
                          ignoreFolders As Variant, allRows As Collection)

    Dim subFolder As Object
    Dim file      As Object

    For Each file In folder.Files
        If Left(file.Name, 1) <> "~" Then
            ' 拡張子が .xlsx であること、かつプレフィックス前方一致でマッチ
            If LCase$(Right$(file.Name, 5)) = ".xlsx" Then
                If Left(NormalizePrefix(file.Name), Len(normalizedTarget)) = normalizedTarget Then
                    Call ReadMainFile(file.Path, srcSheetName, allRows)
                End If
            End If
        End If
    Next file

    For Each subFolder In folder.SubFolders
        If Not IsIgnoredFolder(NormalizeFolderName(subFolder.Name), ignoreFolders) Then
            ' アクセス不可フォルダ（シンボリックリンク等）をスキップ
            If fso.FolderExists(subFolder.Path) Then
                On Error Resume Next
                Call ProcessFolder(fso, subFolder, normalizedTarget, srcSheetName, ignoreFolders, allRows)
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

' ファイル名プレフィックス用の正規化（拡張子除去なし）
' SRC_FILENAME のように "3.2_～" という形式を正しく扱うため
Private Function NormalizePrefix(ByVal name As String) As String

    Dim result As String
    result = LCase$(Trim$(name))
    result = Replace(result, " ", "")
    result = Replace(result, ChrW(&H3000), "")
    result = Replace(result, vbTab, "")
    NormalizePrefix = result

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
    With wsSummary
        .Cells(1, 1).Value = "ソフトウェアユニット仕様ID"
        .Cells(1, 2).Value = "ソフトウェアユニット仕様名"
        .Cells(1, 3).Value = "仕様"
        .Cells(1, 4).Value = "コンポーネント仕様ID一覧"
        .Cells(1, 5).Value = "搭載ソフトウェアユニットID一覧"
        .Cells(1, 6).Value = "搭載ソフトウェアユニット名称一覧"
    End With

    ' ---- 詳細シートを配列に一括読み込み ----
    ' 詳細: col1=参照元識別子, col2=コンポ仕様ID, col5=仕様, col6=ユニット仕様ID,
    '        col7=ユニット仕様名, col10=搭載ユニットID, col11=搭載ユニット名称
    lastRow = wsDetail.Cells(wsDetail.Rows.Count, 6).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim detArr As Variant
    detArr = wsDetail.Range(wsDetail.Cells(2, 1), wsDetail.Cells(lastRow, 11)).Value

    ' ---- メモリ上で集約 ----
    Dim totalRows As Long
    totalRows = UBound(detArr, 1)

    ' 最大でもtotalRows行にはならないが余裕を持って確保
    Dim sumArr() As Variant
    ReDim sumArr(1 To totalRows, 1 To 6)

    Dim outRow  As Long
    outRow = 0
    Dim i As Long
    i = 1

    Do While i <= totalRows

        Dim curUnitID   As String
        Dim curUnitName As String
        Dim curStatus   As String
        Dim funcIDs     As String
        Dim funcNames   As String
        Dim funcCarried As String

        curUnitID   = Trim(CStr(detArr(i, 6)))
        curUnitName = Trim(CStr(detArr(i, 7)))
        curStatus   = Trim(CStr(detArr(i, 5)))
        funcIDs     = ""
        funcNames   = ""
        funcCarried = ""

        Do While i <= totalRows And Trim(CStr(detArr(i, 6))) = curUnitID
            Dim fID      As String
            Dim fCarrID  As String
            Dim fCarrNm  As String
            fID     = Trim(CStr(detArr(i, 2)))
            fCarrID = Trim(CStr(detArr(i, 10)))
            fCarrNm = Trim(CStr(detArr(i, 11)))

            If fID <> "" Then
                If funcIDs = "" Then
                    funcIDs     = fID
                    funcNames   = fCarrID
                    funcCarried = fCarrNm
                Else
                    funcIDs     = funcIDs     & Chr(10) & fID
                    funcNames   = funcNames   & Chr(10) & fCarrID
                    funcCarried = funcCarried & Chr(10) & fCarrNm
                End If
            End If
            i = i + 1
        Loop

        outRow = outRow + 1
        sumArr(outRow, 1) = curUnitID
        sumArr(outRow, 2) = curUnitName
        sumArr(outRow, 3) = curStatus
        sumArr(outRow, 4) = funcIDs
        sumArr(outRow, 5) = funcNames
        sumArr(outRow, 6) = funcCarried

    Loop

    ' ---- 一括書き込み ----
    If outRow > 0 Then
        wsSummary.Range(wsSummary.Cells(2, 1), wsSummary.Cells(1 + outRow, 6)).Value = sumArr
    End If

    ' ---- 折り返し・列幅（ループ外で一括）----
    If outRow > 0 Then
        With wsSummary.Range(wsSummary.Cells(2, 4), wsSummary.Cells(1 + outRow, 6))
            .WrapText = True
        End With
    End If

    wsSummary.Columns("A:F").AutoFit

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
    If outRow > 0 Then
        With wsSummary.Range(wsSummary.Cells(1, 1), wsSummary.Cells(1 + outRow, lastHeaderCol)).Borders
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
