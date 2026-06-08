Option Explicit

' ----------------------------------------------------------
' 「ID一覧」シートを外部ファイルからコピーして更新する
' ----------------------------------------------------------
Public Sub UpdateIDList()

    Const ID_LIST_SHEET As String = "ID一覧"
    Const SRC_FILE_PATH As String = "XXX"   ' コピー元ファイルパスをここに設定
    Const INSERT_AFTER_SHEET As String = "現設計➡"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(SRC_FILE_PATH) Then
        MsgBox "コピー元ファイルが見つかりません:" & vbCrLf & SRC_FILE_PATH, vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' ---- 既存の「ID一覧」シートを削除 ----
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = ID_LIST_SHEET Then
            ws.Delete
            Exit For
        End If
    Next ws

    ' ---- コピー元ファイルを開いてシートをコピー ----
    Dim wbSrc As Workbook
    Set wbSrc = Workbooks.Open(SRC_FILE_PATH, ReadOnly:=True, UpdateLinks:=False)

    Dim wsSrc As Worksheet
    Set wsSrc = Nothing
    For Each ws In wbSrc.Worksheets
        If ws.Name = ID_LIST_SHEET Then
            Set wsSrc = ws
            Exit For
        End If
    Next ws

    If wsSrc Is Nothing Then
        wbSrc.Close SaveChanges:=False
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        MsgBox "コピー元ファイルに「" & ID_LIST_SHEET & "」シートが見つかりません。", vbCritical
        Exit Sub
    End If

    ' ---- 「現設計➡」シートの右側にコピー ----
    Dim wsInsertAfter As Worksheet
    Set wsInsertAfter = Nothing
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = INSERT_AFTER_SHEET Then
            Set wsInsertAfter = ws
            Exit For
        End If
    Next ws

    If wsInsertAfter Is Nothing Then
        wbSrc.Close SaveChanges:=False
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        MsgBox "挿入先シート「" & INSERT_AFTER_SHEET & "」が見つかりません。", vbCritical
        Exit Sub
    End If

    wsSrc.Copy After:=wsInsertAfter

    wbSrc.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "「" & ID_LIST_SHEET & "」シートを更新しました。", vbInformation

End Sub

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
        ' 2行目以降のデータのみ削除（1行目ヘッダーは保持）
        Dim lastDataRow As Long
        lastDataRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row
        If lastDataRow >= 2 Then
            wsTarget.Rows("2:" & lastDataRow).Delete
        End If
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
    ' 重複読み込み防止用（正規化済みパスを記録）
    Dim processedFiles As Object
    Set processedFiles = CreateObject("Scripting.Dictionary")
    ' プレフィックスは拡張子除去不要のため NormalizePrefix を使う
    Call ProcessFolder(fso, fso.GetFolder(unitListFolder), _
                       NormalizePrefix(SRC_FILENAME), SRC_SHEET, ignoreFolders, allRows, processedFiles)

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

    MsgBox "インポート完了。" & (outputRow - 2) & " 件のユニット仕様を書き込みました。", vbInformation

End Sub

' ボタンから呼び出す用エントリーポイント
Public Sub CreateSpecSheet()

    Const UNIT_LIST_SHEET As String = "ユニット仕様一覧"

    Dim wsDetail As Worksheet
    Dim ws       As Worksheet

    Set wsDetail = Nothing
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = UNIT_LIST_SHEET Then
            Set wsDetail = ws
            Exit For
        End If
    Next ws

    If wsDetail Is Nothing Then
        MsgBox "「" & UNIT_LIST_SHEET & "」シートが見つかりません。先にインポートを実行してください。", vbCritical
        Exit Sub
    End If

    Call BuildSpecSheet(wsDetail)
    Call BuildProgressSheet

    MsgBox "「仕様」シートの生成が完了しました。", vbInformation

End Sub

' ----------------------------------------------------------
' 「サマリー」シートを先頭に作成し、仕様シートの関数紐づけ進捗を集計する
' 集計列: コンポーネント名 / ユニット仕様数 / 紐づけ済 / 未紐づけ / 進捗率 / 進捗バー
' ----------------------------------------------------------
Private Sub BuildProgressSheet()

    Const SUMMARY_SHEET As String = "サマリー"
    Const SPEC_SHEET    As String = "仕様"
    Const BAR_MAX       As Integer = 20   ' 進捗バーの最大ブロック数

    Dim wsSum  As Worksheet
    Dim wsSpec As Worksheet
    Dim ws     As Worksheet

    ' ---- 仕様シートの取得 ----
    Set wsSpec = Nothing
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SPEC_SHEET Then
            Set wsSpec = ws
            Exit For
        End If
    Next ws

    If wsSpec Is Nothing Then
        MsgBox "「仕様」シートが見つかりません。", vbCritical
        Exit Sub
    End If

    ' ---- サマリーシートの準備（先頭に配置）----
    Set wsSum = Nothing
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SUMMARY_SHEET Then
            Set wsSum = ws
            Exit For
        End If
    Next ws

    If wsSum Is Nothing Then
        Set wsSum = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        wsSum.Name = SUMMARY_SHEET
    Else
        wsSum.Cells.Clear
        ' 先頭でなければ移動
        If wsSum.Index <> 1 Then
            wsSum.Move Before:=ThisWorkbook.Worksheets(1)
        End If
    End If

    ' ---- 仕様シートを配列に読み込み（2行目以降: A=コンポ名, D=関数名）----
    Dim specLastRow As Long
    specLastRow = wsSpec.Cells(wsSpec.Rows.Count, 2).End(xlUp).Row
    If specLastRow < 2 Then Exit Sub
    
    Dim specArr As Variant
    specArr = wsSpec.Range(wsSpec.Cells(2, 1), wsSpec.Cells(specLastRow, 4)).Value

    ' ---- コンポーネント名ごとに集計（Dictionary使用）----
    Dim dic     As Object
    Dim dicOrd  As Object   ' 出現順を保持
    Set dic    = CreateObject("Scripting.Dictionary")
    Set dicOrd = CreateObject("Scripting.Dictionary")

    Dim i       As Long
    Dim compNm  As String
    Dim funcNm  As String
    Dim ordIdx  As Long
    ordIdx = 0

    For i = 1 To UBound(specArr, 1)
        compNm = Trim(CStr(specArr(i, 1)))   ' A列 = コンポーネント名
        funcNm = Trim(CStr(specArr(i, 4)))   ' D列 = 関数名

        If Not dic.Exists(compNm) Then
            dic.Add compNm, Array(0, 0)   ' (総数, 紐づけ済)
            dicOrd.Add ordIdx, compNm
            ordIdx = ordIdx + 1
        End If

        Dim arr As Variant
        arr = dic(compNm)
        arr(0) = arr(0) + 1
        If funcNm <> "" Then arr(1) = arr(1) + 1
        dic(compNm) = arr
    Next i

    ' ---- タイトル行 ----
    With wsSum
        .Cells(1, 1).Value = "コンポーネント名"
        .Cells(1, 2).Value = "ユニット仕様数"
        .Cells(1, 3).Value = "紐づけ済"
        .Cells(1, 4).Value = "未紐づけ"
        .Cells(1, 5).Value = "進捗率"
        .Cells(1, 6).Value = "進捗"
    End With

    ' ---- コンポーネントごとの行 ----
    Dim outRow As Long
    outRow = 2
    Dim k       As Long
    Dim total   As Long
    Dim done    As Long
    Dim pct     As Double
    Dim bars    As Integer

    For k = 0 To dicOrd.Count - 1
        compNm = dicOrd(k)
        arr    = dic(compNm)
        total  = arr(0)
        done   = arr(1)
        pct    = IIf(total > 0, done / total, 0)
        bars   = Int(pct * BAR_MAX)

        With wsSum
            .Cells(outRow, 1).Value = compNm
            .Cells(outRow, 2).Value = total
            .Cells(outRow, 3).Value = done
            .Cells(outRow, 4).Value = total - done
            .Cells(outRow, 5).Value = pct
            .Cells(outRow, 5).NumberFormat = "0.0%"
            .Cells(outRow, 6).Value = String(bars, ChrW(&H25A0)) & String(BAR_MAX - bars, ChrW(&H25A1))
        End With

        ' 進捗率に応じた色付け
        Dim pctCell As Range
        Set pctCell = wsSum.Cells(outRow, 5)
        If pct >= 1 Then
            pctCell.Interior.Color = RGB(84, 180, 84)    ' 緑（完了）
            pctCell.Font.Color = RGB(255, 255, 255)
        ElseIf pct >= 0.5 Then
            pctCell.Interior.Color = RGB(255, 192, 0)    ' 黄（途中）
            pctCell.Font.Color = RGB(0, 0, 0)
        Else
            pctCell.Interior.Color = RGB(220, 80, 80)    ' 赤（低進捗）
            pctCell.Font.Color = RGB(255, 255, 255)
        End If

        outRow = outRow + 1
    Next k

    ' ---- 合計行 ----
    Dim grandTotal As Long
    Dim grandDone  As Long
    For k = 0 To dicOrd.Count - 1
        arr = dic(dicOrd(k))
        grandTotal = grandTotal + arr(0)
        grandDone  = grandDone  + arr(1)
    Next k

    Dim grandPct As Double
    grandPct = IIf(grandTotal > 0, grandDone / grandTotal, 0)
    bars = Int(grandPct * BAR_MAX)

    With wsSum
        .Cells(outRow, 1).Value = "【合計】"
        .Cells(outRow, 2).Value = grandTotal
        .Cells(outRow, 3).Value = grandDone
        .Cells(outRow, 4).Value = grandTotal - grandDone
        .Cells(outRow, 5).Value = grandPct
        .Cells(outRow, 5).NumberFormat = "0.0%"
        .Cells(outRow, 6).Value = String(bars, ChrW(&H25A0)) & String(BAR_MAX - bars, ChrW(&H25A1))

        ' 合計行の背景
        .Range(.Cells(outRow, 1), .Cells(outRow, 6)).Interior.Color = RGB(30, 80, 160)
        .Range(.Cells(outRow, 1), .Cells(outRow, 6)).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(outRow, 1), .Cells(outRow, 6)).Font.Bold = True
        .Cells(outRow, 5).Interior.Color = RGB(30, 80, 160)
    End With

    ' ---- ヘッダー書式 ----
    With wsSum.Range(wsSum.Cells(1, 1), wsSum.Cells(1, 6))
        .Interior.Color = RGB(0, 112, 192)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    ' ---- 罫線 ----
    With wsSum.Range(wsSum.Cells(1, 1), wsSum.Cells(outRow, 6)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(180, 180, 180)
    End With

    ' ---- 列幅 ----
    wsSum.Columns("A:E").AutoFit
    wsSum.Columns("F").ColumnWidth = 40

    ' F2:F16 フォントサイズ10
    wsSum.Range("F2:F16").Font.Size = 10

End Sub

Private Sub ProcessFolder(fso As Object, folder As Object, _
                          normalizedTarget As String, srcSheetName As String, _
                          ignoreFolders As Variant, allRows As Collection, _
                          processedFiles As Object)

    Dim subFolder As Object
    Dim file      As Object

    For Each file In folder.Files
        If Left(file.Name, 1) <> "~" Then
            ' 拡張子が .xlsx であること、かつプレフィックス前方一致でマッチ
            If LCase$(Right$(file.Name, 5)) = ".xlsx" Then
                If Left(NormalizePrefix(file.Name), Len(normalizedTarget)) = normalizedTarget Then
                    ' 重複読み込みをスキップ
                    Dim normalizedPath As String
                    normalizedPath = LCase$(Trim$(file.Path))
                    If Not processedFiles.Exists(normalizedPath) Then
                        processedFiles.Add normalizedPath, True
                        Call ReadMainFile(file.Path, srcSheetName, allRows)
                    End If
                End If
            End If
        End If
    Next file

    For Each subFolder In folder.SubFolders
        If Not IsIgnoredFolder(NormalizeFolderName(subFolder.Name), ignoreFolders) Then
            ' アクセス不可フォルダ（シンボリックリンク等）をスキップ
            If fso.FolderExists(subFolder.Path) Then
                On Error Resume Next
                Call ProcessFolder(fso, subFolder, normalizedTarget, srcSheetName, ignoreFolders, allRows, processedFiles)
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

' ----------------------------------------------------------
' 「仕様」シートを「ユニット仕様一覧」と「ID一覧」から生成する
'
' 出力列:
'   A: 機能名         (ID一覧 C列 ← ID一覧 A列 = ユニット仕様一覧 F列)
'   B: ソフトウェアユニット仕様ID   (ユニット仕様一覧 F列)
'   C: ソフトウェアユニット仕様     (ユニット仕様一覧 G列)
'   D: 関数名         (ID一覧 I列 ← ID一覧 A列 = B列 AND ID一覧 D列 = C列)
' ----------------------------------------------------------
Private Sub BuildSpecSheet(wsDetail As Worksheet)

    Const SPEC_SHEET    As String = "仕様"
    Const FUNC_SHEET    As String = "関数"
    Const FUNC_LIST_SHEET As String = "関数一覧"
    Const ID_LIST_SHEET As String = "ID一覧"
    Const ID_HEADER_ROW As Long   = 6  ' ID一覧のヘッダー行

    Dim wsSpec   As Worksheet
    Dim wsFunc   As Worksheet
    Dim wsFuncList As Worksheet
    Dim wsIDList As Worksheet
    Dim ws       As Worksheet

    ' ---- 仕様シートの取得 ----
    Set wsSpec = Nothing
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SPEC_SHEET Then
            Set wsSpec = ws
            Exit For
        End If
    Next ws

    If wsSpec Is Nothing Then
        MsgBox "「仕様」シートが見つかりません。", vbCritical
        Exit Sub
    End If

    ' ---- ID一覧シートの取得 ----
    Set wsIDList = Nothing
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = ID_LIST_SHEET Then
            Set wsIDList = ws
            Exit For
        End If
    Next ws

    If wsIDList Is Nothing Then
        MsgBox "「ID一覧」シートが見つかりません。", vbCritical
        Exit Sub
    End If

    ' ---- 関数シートの取得 ----
    Set wsFunc = Nothing
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = FUNC_SHEET Then
            Set wsFunc = ws
            Exit For
        End If
    Next ws

    ' ---- 関数一覧シートの取得 ----
    Set wsFuncList = Nothing
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = FUNC_LIST_SHEET Then
            Set wsFuncList = ws
            Exit For
        End If
    Next ws

    ' ---- 仕様シートのクリア ----
    wsSpec.Cells.Clear

    ' ---- 関数シートの2行目以降をクリア ----
    If Not wsFunc Is Nothing Then
        Dim funcLastRow As Long
        funcLastRow = wsFunc.Cells(wsFunc.Rows.Count, 2).End(xlUp).Row
        If funcLastRow >= 2 Then
            wsFunc.Rows("2:" & funcLastRow).Delete
        End If
    End If

    ' ---- タイトル行 ----
    With wsSpec
        .Cells(1, 1).Value = "コンポーネント名"
        .Cells(1, 2).Value = "ソフトウェアユニット仕様ID"
        .Cells(1, 3).Value = "ソフトウェアユニット仕様"
        .Cells(1, 4).Value = "関数名"
    End With

    ' ---- ユニット仕様一覧を配列に読み込み ----
    ' col1=参照元識別子, col2=コンポーネント仕様ID, col6=ユニット仕様ID, col7=ユニット仕様
    Dim detLastRow As Long
    detLastRow = wsDetail.Cells(wsDetail.Rows.Count, 6).End(xlUp).Row
    If detLastRow < 2 Then Exit Sub

    Dim detArr As Variant
    detArr = wsDetail.Range(wsDetail.Cells(2, 1), wsDetail.Cells(detLastRow, 11)).Value

    ' ---- ID一覧を配列に読み込み（ヘッダー行の次行から）----
    ' col1=ソフトウェアユニット仕様ID, col3=機能名, col4=ソフトウェアユニット仕様, col9=関数名
    Dim idLastRow As Long
    idLastRow = wsIDList.Cells(wsIDList.Rows.Count, 1).End(xlUp).Row
    If idLastRow <= ID_HEADER_ROW Then Exit Sub

    Dim idArr As Variant
    idArr = wsIDList.Range(wsIDList.Cells(ID_HEADER_ROW + 1, 1), _
                           wsIDList.Cells(idLastRow, 9)).Value

    ' ---- 出力配列の構築 ----
    Dim totalRows As Long
    totalRows = UBound(detArr, 1)

    Dim outArr() As Variant
    ReDim outArr(1 To totalRows, 1 To 4)

    Dim i As Long, j As Long
    Dim compName    As String   ' ユニット仕様一覧 A列（参照元識別子）
    Dim unitID      As String
    Dim unitSpec    As String
    Dim funcName    As String
    Dim idA         As String
    Dim idD         As String

    For i = 1 To totalRows
        compName = Trim(CStr(detArr(i, 1)))   ' ユニット仕様一覧 A列（参照元識別子）
        unitID   = Trim(CStr(detArr(i, 6)))   ' ユニット仕様一覧 F列
        unitSpec = Trim(CStr(detArr(i, 7)))   ' ユニット仕様一覧 G列
        funcName    = ""

        For j = 1 To UBound(idArr, 1)
            idA = Trim(CStr(idArr(j, 1)))   ' ID一覧 A列
            idD = Trim(CStr(idArr(j, 4)))   ' ID一覧 D列
 
            If idA = unitID And idD = unitSpec Then
                ' 関数名: 一致する全行を改行で結合
                Dim fn As String
                fn = Trim(CStr(idArr(j, 9)))   ' ID一覧 I列
                If fn <> "" Then
                    If funcName = "" Then
                        funcName = fn
                    Else
                        funcName = funcName & Chr(10) & fn
                    End If
                End If
            End If
        Next j

        outArr(i, 1) = compName     ' 仕様 A列 = コンポーネント名（参照元識別子）
        outArr(i, 2) = unitID       ' 仕様 B列 = ソフトウェアユニット仕様ID
        outArr(i, 3) = unitSpec     ' 仕様 C列 = ソフトウェアユニット仕様
        outArr(i, 4) = funcName     ' 仕様 D列 = 関数名
    Next i

    ' ---- 一括書き込み（2行目から）----
    wsSpec.Range(wsSpec.Cells(2, 1), wsSpec.Cells(1 + totalRows, 4)).Value = outArr

    ' ---- 関数シートへ書き込み（B列から、2行目以降）----
    If Not wsFunc Is Nothing Then
        wsFunc.Range(wsFunc.Cells(2, 2), wsFunc.Cells(1 + totalRows, 5)).Value = outArr
        FillExistingFunctionNames wsFunc, wsFuncList, totalRows
    End If

    ' ---- 列幅自動調整・折り返し（関数名列）----
    wsSpec.Columns("A:D").AutoFit
    wsSpec.Columns("D").WrapText = True
    If Not wsFunc Is Nothing Then
        wsFunc.Columns("E").WrapText = True
    End If

    ' ---- タイトル行の書式（A1:D1）----
    With wsSpec.Range("A1:D1")
        .Interior.Color = RGB(0, 112, 192)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    ' ---- 罫線（データ全体）----
    Dim specDataLastRow As Long
    specDataLastRow = wsSpec.Cells(wsSpec.Rows.Count, 2).End(xlUp).Row
    If specDataLastRow >= 1 Then
        With wsSpec.Range(wsSpec.Cells(1, 1), wsSpec.Cells(specDataLastRow, 4)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
    End If

End Sub

Private Sub FillExistingFunctionNames(ByVal wsFunc As Worksheet, ByVal wsFuncList As Worksheet, ByVal totalRows As Long)

    If wsFunc Is Nothing Then Exit Sub
    If wsFuncList Is Nothing Then Exit Sub
    If totalRows <= 0 Then Exit Sub

    Dim funcListLastRow As Long
    funcListLastRow = wsFuncList.Cells(wsFuncList.Rows.Count, 1).End(xlUp).Row
    If funcListLastRow < 2 Then Exit Sub

    Dim listArr As Variant
    listArr = wsFuncList.Range(wsFuncList.Cells(2, 1), wsFuncList.Cells(funcListLastRow, 1)).Value

    Dim funcDic As Object
    Set funcDic = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim funcName As String
    For i = 1 To UBound(listArr, 1)
        funcName = Trim(CStr(listArr(i, 1)))
        If funcName <> "" Then
            If Not funcDic.Exists(funcName) Then
                funcDic.Add funcName, funcName
            End If
        End If
    Next i

    If funcDic.Count = 0 Then Exit Sub

    Dim funcArr As Variant
    funcArr = wsFunc.Range(wsFunc.Cells(2, 5), wsFunc.Cells(1 + totalRows, 5)).Value

    Dim existingArr() As Variant
    ReDim existingArr(1 To totalRows, 1 To 1)

    For i = 1 To totalRows
        existingArr(i, 1) = FindMatchingFunctions(CStr(funcArr(i, 1)), funcDic)
    Next i

    wsFunc.Range(wsFunc.Cells(2, 1), wsFunc.Cells(1 + totalRows, 1)).Value = existingArr
    wsFunc.Columns("A").WrapText = True

End Sub

Private Function FindMatchingFunctions(ByVal functionText As String, ByVal funcDic As Object) As String

    Dim normalizedText As String
    normalizedText = Replace(functionText, vbCrLf, vbLf)
    normalizedText = Replace(normalizedText, vbCr, vbLf)

    Dim items() As String
    items = Split(normalizedText, vbLf)

    Dim i As Long
    Dim item As String
    Dim matches As String

    For i = LBound(items) To UBound(items)
        item = Trim$(items(i))
        If item <> "" Then
            If funcDic.Exists(item) Then
                If matches = "" Then
                    matches = item
                ElseIf InStr(1, vbLf & matches & vbLf, vbLf & item & vbLf, vbBinaryCompare) = 0 Then
                    matches = matches & vbLf & item
                End If
            End If
        End If
    Next i

    FindMatchingFunctions = matches

End Function

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
