Attribute VB_Name = "ふるさと納税_マスター"
'==================================================
' ふるさと納税(小牧市) 出荷自動化 - マスター実行モジュール
' Amazon/楽天/ペライチと同じ「実行1/2/3ボタン」構成。
'
' 【ボタン配置】
'   実行1_ふるさと納税_編集データ作成 → 原本→編集＋発送方法(クリック/佐川)自動判定
'   実行2_ふるさと納税_出荷データ作成  → 振り分け＋物出しリスト＋納品書を一括作成
'   実行3_ふるさと納税_出荷通知        → ポータルへの出荷通知（仕様確定後に実装）
'
' 【他チャネルとの違い】
'   ・謝礼品はすべてペット用品のため「衣類判定」は無い。発送方法はクリック/佐川の2択。
'   ・送り主は「株式会社SEED(shop fuu)」で固定（※旧テンプレのARCS・古雅61-4は誤り。
'     正しくは株式会社SEED／古雅3丁目61番3。SEED納品書フォーマットに準拠）。
'   ・クリック/e飛伝3へは既存数式に頼らず値で直接転記（データがずれても壊れにくい）。
'   ・物出しリスト・納品書シートは無ければ自動作成する。
'
' 抽出元: 各種フォーマット\★ふるさと納税出荷　フォーマット.xlsx
'==================================================

'--- 送り主（株式会社SEED）固定情報 ------------------------
Private Const SEED_社名 As String = "株式会社ＳＥＥＤ"
Private Const SEED_屋号 As String = "shop fuu ／ 株式会社ＳＥＥＤ"
Private Const SEED_郵便 As String = "485-0814"
Private Const SEED_住所県 As String = "愛知県小牧市"
Private Const SEED_住所番地 As String = "古雅３－６１－３"
Private Const SEED_住所表示 As String = "愛知県小牧市古雅3丁目61番3"
Private Const SEED_電話 As String = "0568-47-5090"
Private Const 内容品 As String = "小牧市ふるさと納税謝礼品"

'--- 「原本」シートの列位置（1始まり） ---------------------
Private Const 原_配送No As Long = 2      ' B
Private Const 原_謝礼品 As Long = 4      ' D
Private Const 原_配送姓 As Long = 7      ' G
Private Const 原_配送名 As Long = 8      ' H
Private Const 原_都道府県 As Long = 11   ' K
Private Const 原_郵便番号 As Long = 12   ' L
Private Const 原_住所1 As Long = 13      ' M
Private Const 原_住所2 As Long = 14      ' N
Private Const 原_電話 As Long = 15       ' O

'--- 「編集」シートは原本を1列右へずらしてコピー（編集列 = 原本列 + 1）。A列=発送方法 ---
Private Const 編_発送方法 As Long = 1    ' A（新規）
Private Const 編_配送No As Long = 原_配送No + 1     ' C
Private Const 編_謝礼品 As Long = 原_謝礼品 + 1     ' E
Private Const 編_配送姓 As Long = 原_配送姓 + 1     ' H
Private Const 編_配送名 As Long = 原_配送名 + 1     ' I
Private Const 編_都道府県 As Long = 原_都道府県 + 1 ' L
Private Const 編_郵便番号 As Long = 原_郵便番号 + 1 ' M
Private Const 編_住所1 As Long = 原_住所1 + 1       ' N
Private Const 編_住所2 As Long = 原_住所2 + 1       ' O
Private Const 編_電話 As Long = 原_電話 + 1         ' P


'==================================================
' 実行1: 原本 → 編集シート作成 + 発送方法(クリック/佐川)を自動判定
'==================================================
Sub 実行1_ふるさと納税_編集データ作成()
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    Dim wsGen As Worksheet, wsEdit As Worksheet
    Set wsGen = Sheets("原本")
    Set wsEdit = GetOrCreateSheet("編集")

    ' 編集シートを作り直す（原本を値貼り → A列を1列挿入して発送方法欄に）
    wsEdit.Cells.Clear
    wsGen.Cells.Copy
    wsEdit.Range("A1").PasteSpecial Paste:=xlPasteValues
    wsEdit.Columns("A:A").Insert Shift:=xlToRight
    Application.CutCopyMode = False
    wsEdit.Cells(1, 編_発送方法).Value = "発送方法"
    wsEdit.Columns(編_発送方法).ColumnWidth = 10

    Dim lastRow As Long
    lastRow = wsEdit.Cells(wsEdit.Rows.Count, 編_配送姓).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "原本に処理対象のデータがありません。", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' A列に発送方法を自動判定（謝礼品名の重量・サイズで佐川/クリックを振り分け）
    Dim i As Long, gift As String
    For i = 2 To lastRow
        gift = CStr(wsEdit.Cells(i, 編_謝礼品).Value)
        wsEdit.Cells(i, 編_発送方法).Value = IIf(Is佐川対象(gift), "佐川", "クリック")
    Next i

    ' 自動判定を手で直せるよう、A列にドロップダウンを付ける
    With wsEdit.Range(wsEdit.Cells(2, 編_発送方法), wsEdit.Cells(lastRow, 編_発送方法)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="クリック,佐川"
        .InCellDropdown = True
    End With

    Dim nClick As Long, nSagawa As Long
    nClick = Application.WorksheetFunction.CountIf( _
        wsEdit.Range(wsEdit.Cells(2, 編_発送方法), wsEdit.Cells(lastRow, 編_発送方法)), "クリック")
    nSagawa = (lastRow - 1) - nClick

    Application.ScreenUpdating = True
    MsgBox "編集データを作成しました（" & (lastRow - 1) & "件）。" & vbCrLf & vbCrLf & _
           "　クリックポスト: " & nClick & "件" & vbCrLf & _
           "　佐川: " & nSagawa & "件" & vbCrLf & vbCrLf & _
           "「編集」シートのA列(発送方法)を確認し、必要なら直してから" & vbCrLf & _
           "「実行2」を押してください。", vbInformation, "実行1 完了"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "処理中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical, "エラー"
End Sub


'==================================================
' 実行2: 出荷データ作成一式（振り分け → 物出しリスト → 納品書）
'==================================================
Sub 実行2_ふるさと納税_出荷データ作成()
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    Dim wsEdit As Worksheet
    Set wsEdit = Sheets("編集")

    Dim lastRow As Long
    lastRow = wsEdit.Cells(wsEdit.Rows.Count, 編_配送姓).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "編集シートにデータがありません。先に実行1を行ってください。", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    Call 振り分け_発行用(wsEdit, lastRow)
    Call 物出しリスト作成(wsEdit, lastRow)
    Call 納品書作成(wsEdit, lastRow)

    Application.ScreenUpdating = True
    MsgBox "出荷データ一式を作成しました。" & vbCrLf & vbCrLf & _
           "　・クリック / e飛伝3 … 宛名・伝票の発行用データ" & vbCrLf & _
           "　・物出しリスト … ピッキング用（発送方法ごとに並べ替え済み）" & vbCrLf & _
           "　・納品書 … 寄附者ごと（1ページ1名、印刷してそのまま同梱可）" & vbCrLf & vbCrLf & _
           "1. クリックポスト/e飛伝で宛名・伝票を発行" & vbCrLf & _
           "2. 物出しリストで商品を用意、納品書を印刷して同梱" & vbCrLf & _
           "3. 伝票番号を控えたら「実行3」へ", vbInformation, "実行2 完了"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "処理中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical, "エラー"
End Sub


'--------------------------------------------------
' 振り分け: クリック宛→「クリック」、佐川宛→「e飛伝3」へ値で転記（送り主SEED固定）
'--------------------------------------------------
Private Sub 振り分け_発行用(wsEdit As Worksheet, lastRow As Long)
    Dim wsClick As Worksheet, wsHiden As Worksheet
    Set wsClick = Sheets("クリック")
    Set wsHiden = Sheets("e飛伝3")

    ClearDataRows wsClick, 1, 9
    ClearDataRows wsHiden, 1, 40

    Dim rClick As Long: rClick = 2
    Dim rHiden As Long: rHiden = 2

    Dim i As Long, method As String, fullName As String
    For i = 2 To lastRow
        method = CStr(wsEdit.Cells(i, 編_発送方法).Value)
        fullName = CStr(wsEdit.Cells(i, 編_配送姓).Value) & "　" & CStr(wsEdit.Cells(i, 編_配送名).Value)

        If method = "クリック" Then
            wsClick.Cells(rClick, "B").Value = wsEdit.Cells(i, 編_郵便番号).Value
            wsClick.Cells(rClick, "C").Value = fullName
            wsClick.Cells(rClick, "D").Value = "様"
            wsClick.Cells(rClick, "E").Value = wsEdit.Cells(i, 編_都道府県).Value
            wsClick.Cells(rClick, "F").Value = wsEdit.Cells(i, 編_住所1).Value
            wsClick.Cells(rClick, "G").Value = wsEdit.Cells(i, 編_住所2).Value
            wsClick.Cells(rClick, "I").Value = 内容品
            rClick = rClick + 1
        Else
            wsHiden.Cells(rHiden, "C").Value = wsEdit.Cells(i, 編_電話).Value
            wsHiden.Cells(rHiden, "D").Value = wsEdit.Cells(i, 編_郵便番号).Value
            wsHiden.Cells(rHiden, "E").Value = wsEdit.Cells(i, 編_都道府県).Value
            wsHiden.Cells(rHiden, "F").Value = wsEdit.Cells(i, 編_住所1).Value
            wsHiden.Cells(rHiden, "G").Value = wsEdit.Cells(i, 編_住所2).Value
            wsHiden.Cells(rHiden, "H").Value = fullName
            ' 送り主（株式会社SEED）固定
            wsHiden.Cells(rHiden, "R").Value = SEED_電話
            wsHiden.Cells(rHiden, "S").Value = SEED_郵便
            wsHiden.Cells(rHiden, "T").Value = SEED_住所県
            wsHiden.Cells(rHiden, "U").Value = SEED_住所番地
            wsHiden.Cells(rHiden, "V").Value = SEED_社名
            ' 品名
            wsHiden.Cells(rHiden, "Y").Value = "小牧市ふるさと納税"
            wsHiden.Cells(rHiden, "Z").Value = "謝礼品"
            wsHiden.Cells(rHiden, "AA").Value = "ペット用品"
            rHiden = rHiden + 1
        End If
    Next i
End Sub


'--------------------------------------------------
' 物出しリスト: ピッキング用。発送方法ごとに並べて一覧化
'--------------------------------------------------
Private Sub 物出しリスト作成(wsEdit As Worksheet, lastRow As Long)
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet("物出しリスト")
    ws.Cells.Clear

    ws.Range("A1:F1").Value = Array("№", "発送方法", "お客様名", "〒", "住所", "謝礼品名")
    ws.Range("A1:F1").Font.Bold = True

    ' クリック→佐川の順に並べて書き出す（ピッキングしやすくするため）
    Dim r As Long: r = 2
    Dim pass As Long, wantClick As Boolean
    For pass = 1 To 2
        wantClick = (pass = 1)
        Dim i As Long
        For i = 2 To lastRow
            Dim m As String: m = CStr(wsEdit.Cells(i, 編_発送方法).Value)
            If (wantClick And m = "クリック") Or ((Not wantClick) And m <> "クリック") Then
                ws.Cells(r, 1).Value = r - 1
                ws.Cells(r, 2).Value = m
                ws.Cells(r, 3).Value = CStr(wsEdit.Cells(i, 編_配送姓).Value) & "　" & CStr(wsEdit.Cells(i, 編_配送名).Value)
                ws.Cells(r, 4).Value = wsEdit.Cells(i, 編_郵便番号).Value
                ws.Cells(r, 5).Value = CStr(wsEdit.Cells(i, 編_都道府県).Value) & CStr(wsEdit.Cells(i, 編_住所1).Value) & CStr(wsEdit.Cells(i, 編_住所2).Value)
                ws.Cells(r, 6).Value = wsEdit.Cells(i, 編_謝礼品).Value
                r = r + 1
            End If
        Next i
    Next pass

    ws.Columns("A:F").EntireColumn.AutoFit
End Sub


'--------------------------------------------------
' 納品書: 寄附者ごとに1ページ。SEED社名で印刷してそのまま同梱できる
'--------------------------------------------------
Private Sub 納品書作成(wsEdit As Worksheet, lastRow As Long)
    Const BLOCK_H As Long = 16   ' 1名分の高さ（行）
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet("納品書")
    ws.Cells.Clear
    ws.ResetAllPageBreaks

    Dim blk As Long: blk = 0
    Dim i As Long
    For i = 2 To lastRow
        Dim base As Long: base = 1 + blk * BLOCK_H

        ' 2件目以降は改ページを入れて1名1ページにする
        If blk > 0 Then ws.HPageBreaks.Add Before:=ws.Rows(base)

        Dim fullName As String
        fullName = CStr(wsEdit.Cells(i, 編_配送姓).Value) & "　" & CStr(wsEdit.Cells(i, 編_配送名).Value)

        ' タイトル・発行日
        ws.Cells(base, 1).Value = "納　品　書"
        ws.Cells(base, 1).Font.Size = 16
        ws.Cells(base, 6).Value = "発行日"
        ws.Cells(base, 7).Value = Date
        ws.Cells(base, 7).NumberFormat = "yyyy/mm/dd"

        ' お届け先（左）
        ws.Cells(base + 1, 1).Value = fullName & " 様"
        ws.Cells(base + 1, 1).Font.Size = 12
        ws.Cells(base + 2, 1).Value = "〒" & CStr(wsEdit.Cells(i, 編_郵便番号).Value)
        ws.Cells(base + 3, 1).Value = CStr(wsEdit.Cells(i, 編_都道府県).Value) & CStr(wsEdit.Cells(i, 編_住所1).Value)
        ws.Cells(base + 4, 1).Value = CStr(wsEdit.Cells(i, 編_住所2).Value)

        ' 送り主SEED（右）
        ws.Cells(base + 1, 6).Value = SEED_屋号
        ws.Cells(base + 2, 6).Value = "〒" & SEED_郵便
        ws.Cells(base + 3, 6).Value = SEED_住所表示
        ws.Cells(base + 4, 6).Value = "TEL:" & SEED_電話
        ws.Cells(base + 5, 6).Value = "事業者番号：T5180001078443"

        ' あいさつ
        ws.Cells(base + 7, 1).Value = "この度はふるさと納税にお申込みいただき、誠にありがとうございます。"

        ' 明細（謝礼品は寄附のお礼のため金額表示なし）
        ws.Cells(base + 9, 1).Value = "品名"
        ws.Cells(base + 9, 5).Value = "数量"
        ws.Cells(base + 9, 1).Font.Bold = True
        ws.Cells(base + 9, 5).Font.Bold = True
        ws.Cells(base + 10, 1).Value = 内容品 & "（" & CStr(wsEdit.Cells(i, 編_謝礼品).Value) & "）"
        ws.Cells(base + 10, 5).Value = 1

        blk = blk + 1
    Next i

    ws.Columns("A:G").EntireColumn.AutoFit
    ws.PageSetup.Orientation = xlPortrait
End Sub


'==================================================
' 実行3: 出荷通知（※未確定。ポータルへの通知仕様が分かり次第実装）
'==================================================
Sub 実行3_ふるさと納税_出荷通知()
    MsgBox "【準備中】出荷通知。" & vbCrLf & vbCrLf & _
           "ふるさと納税ポータルへの発送通知で登録する列(伝票番号・配達業者など)と" & vbCrLf & _
           "出力フォーマットが確定してから、他チャネルの手順7/8に倣って実装します。" & vbCrLf & _
           "現在の通知のやり方を教えてください。", vbInformation, "実行3（準備中）"
End Sub


'==================================================
' 補助関数
'==================================================

'--- 謝礼品名から「佐川対象（重い/大型）」かどうかを判定 ---
'   ※キーワードは要確認・要調整。重い謝礼品名に出てくる語をカンマで足し引きするだけで直せます。
Function Is佐川対象(giftName As String) As Boolean
    Dim keywords As Variant, k As Variant
    keywords = Array("3kg", "5kg", "10kg", "3ｋｇ", "5ｋｇ", "10ｋｇ", _
                     "大容量", "大袋", "ケース", "箱")
    Is佐川対象 = False
    For Each k In keywords
        If InStr(giftName, CStr(k)) > 0 Then
            Is佐川対象 = True
            Exit Function
        End If
    Next k
End Function

'--- 指定シートの2行目以降をクリア（ヘッダ行は残す） ---
Private Sub ClearDataRows(ws As Worksheet, headerRow As Long, lastCol As Long)
    ws.Range(ws.Cells(headerRow + 1, 1), ws.Cells(ws.Rows.Count, lastCol)).ClearContents
End Sub

'--- シートを取得（無ければ末尾に新規作成） ---
Private Function GetOrCreateSheet(sheetName As String) As Worksheet
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = sheetName Then
            Set GetOrCreateSheet = sh
            Exit Function
        End If
    Next sh
    Set GetOrCreateSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    GetOrCreateSheet.Name = sheetName
End Function
