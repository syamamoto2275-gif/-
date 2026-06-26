Attribute VB_Name = "ふるさと納税_マスター"
'==================================================
' ふるさと納税(小牧市) 出荷自動化 - マスター実行モジュール
' ボタン3つで全工程を制御する
'
' 【ボタン配置】
'   実行1_ふるさと納税_編集データ作成 → Sub 実行1_ふるさと納税_編集データ作成
'   実行2_ふるさと納税_出荷データ作成  → Sub 実行2_ふるさと納税_出荷データ作成
'   実行3_ふるさと納税_出荷通知        → Sub 実行3_ふるさと納税_出荷通知
'
' 【他チャネルとの違い】
'   ・謝礼品はすべてペット用品のため「衣類判定」は無い。
'   ・発送方法は「クリック」か「佐川」かの2択（ペット/衣類の区別なし）。
'   ・送り主は株式会社ARCS(小牧市)で固定。
'   ・クリック/e飛伝3/注文確認表シートへは、既存の数式に頼らず値で直接転記する
'     （データがずれても壊れにくくするため。Amazon/Wowmaと同じ方式）。
'
' 抽出元フォーマット: 各種フォーマット\★ふるさと納税出荷　フォーマット.xlsx
'==================================================

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
Private Const 原_最終列 As Long = 40     ' AN（原本のデータ最終列）

'--- 「編集」シートは原本を1列右へずらしてコピーし、A列に発送方法を持つ ---
'    （編集の各列 = 原本の対応列 + 1）
Private Const 編_発送方法 As Long = 1    ' A（新規）
Private Const 編_謝礼品 As Long = 原_謝礼品 + 1   ' E
Private Const 編_配送姓 As Long = 原_配送姓 + 1   ' H
Private Const 編_配送名 As Long = 原_配送名 + 1   ' I
Private Const 編_都道府県 As Long = 原_都道府県 + 1 ' L
Private Const 編_郵便番号 As Long = 原_郵便番号 + 1 ' M
Private Const 編_住所1 As Long = 原_住所1 + 1     ' N
Private Const 編_住所2 As Long = 原_住所2 + 1     ' O
Private Const 編_電話 As Long = 原_電話 + 1       ' P

Private Const 内容品 As String = "小牧市ふるさと納税謝礼品"


'--------------------------------------------------
' 実行1: 原本 → 編集シート作成 + 発送方法(クリック/佐川)を自動判定
'   ここで一度A列を目視確認・修正してから実行2へ進む
'--------------------------------------------------
Sub 実行1_ふるさと納税_編集データ作成()
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    Dim wsGen As Worksheet, wsEdit As Worksheet
    Set wsGen = Sheets("原本")
    Set wsEdit = Sheets("編集")

    ' 編集シートを作り直す（原本を丸ごとコピーし、A列を1列挿入して発送方法欄にする）
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

    ' 件数を集計して報告
    Dim nClick As Long, nSagawa As Long
    nClick = Application.WorksheetFunction.CountIf( _
        wsEdit.Range(wsEdit.Cells(2, 編_発送方法), wsEdit.Cells(lastRow, 編_発送方法)), "クリック")
    nSagawa = (lastRow - 1) - nClick

    Application.ScreenUpdating = True
    MsgBox "編集データを作成しました（" & (lastRow - 1) & "件）。" & vbCrLf & vbCrLf & _
           "　クリックポスト: " & nClick & "件" & vbCrLf & _
           "　佐川: " & nSagawa & "件" & vbCrLf & vbCrLf & _
           "「編集」シートのA列(発送方法)を確認してください。" & vbCrLf & _
           "重い謝礼品で佐川にすべき／クリックの方が安い行があれば、" & vbCrLf & _
           "ドロップダウンで選び直してから「実行2」を押してください。", _
           vbInformation, "実行1 完了"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "処理中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical, "エラー"
End Sub


'--------------------------------------------------
' 実行2: 編集データを発送方法で振り分けて発行用データを作成
'   クリック宛 → 「クリック」シート、佐川宛 → 「e飛伝3」シートへ値で転記
'--------------------------------------------------
Sub 実行2_ふるさと納税_出荷データ作成()
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    Dim wsEdit As Worksheet, wsClick As Worksheet, wsHiden As Worksheet
    Set wsEdit = Sheets("編集")
    Set wsClick = Sheets("クリック")
    Set wsHiden = Sheets("e飛伝3")

    Dim lastRow As Long
    lastRow = wsEdit.Cells(wsEdit.Rows.Count, 編_配送姓).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "編集シートにデータがありません。先に実行1を行ってください。", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' 各発行用シートのデータ行(2行目以降)をクリア（ヘッダ行1は残す）
    ClearDataRows wsClick, 1, 9
    ClearDataRows wsHiden, 1, 40

    Dim rClick As Long, rHiden As Long
    rClick = 2
    rHiden = 2

    Dim i As Long, method As String
    Dim sei As String, mei As String, fullName As String
    For i = 2 To lastRow
        method = CStr(wsEdit.Cells(i, 編_発送方法).Value)
        sei = CStr(wsEdit.Cells(i, 編_配送姓).Value)
        mei = CStr(wsEdit.Cells(i, 編_配送名).Value)
        fullName = sei & "　" & mei

        If method = "クリック" Then
            ' クリックポスト発行用（A=通知番号は発行後に入る）
            wsClick.Cells(rClick, "B").Value = wsEdit.Cells(i, 編_郵便番号).Value
            wsClick.Cells(rClick, "C").Value = fullName
            wsClick.Cells(rClick, "D").Value = "様"
            wsClick.Cells(rClick, "E").Value = wsEdit.Cells(i, 編_都道府県).Value
            wsClick.Cells(rClick, "F").Value = wsEdit.Cells(i, 編_住所1).Value
            wsClick.Cells(rClick, "G").Value = wsEdit.Cells(i, 編_住所2).Value
            wsClick.Cells(rClick, "I").Value = 内容品
            rClick = rClick + 1
        Else
            ' 佐川 e飛伝発行用
            wsHiden.Cells(rHiden, "C").Value = wsEdit.Cells(i, 編_電話).Value
            wsHiden.Cells(rHiden, "D").Value = wsEdit.Cells(i, 編_郵便番号).Value
            wsHiden.Cells(rHiden, "E").Value = wsEdit.Cells(i, 編_都道府県).Value
            wsHiden.Cells(rHiden, "F").Value = wsEdit.Cells(i, 編_住所1).Value
            wsHiden.Cells(rHiden, "G").Value = wsEdit.Cells(i, 編_住所2).Value
            wsHiden.Cells(rHiden, "H").Value = fullName
            ' 送り主（株式会社ARCS・小牧市）固定
            wsHiden.Cells(rHiden, "R").Value = "0568-47-5090"
            wsHiden.Cells(rHiden, "S").Value = "485-0814"
            wsHiden.Cells(rHiden, "T").Value = "愛知県小牧市"
            wsHiden.Cells(rHiden, "U").Value = "古雅３－６１－４"
            wsHiden.Cells(rHiden, "V").Value = "株式会社ＡＲＣＳ"
            ' 品名
            wsHiden.Cells(rHiden, "Y").Value = "小牧市ふるさと納税"
            wsHiden.Cells(rHiden, "Z").Value = "謝礼品"
            wsHiden.Cells(rHiden, "AA").Value = "ペット用品"
            rHiden = rHiden + 1
        End If
    Next i

    Application.ScreenUpdating = True
    MsgBox "出荷データを作成しました。" & vbCrLf & vbCrLf & _
           "　クリックポスト発行用: " & (rClick - 2) & "件 →「クリック」シート" & vbCrLf & _
           "　佐川 e飛伝発行用: " & (rHiden - 2) & "件 →「e飛伝3」シート" & vbCrLf & vbCrLf & _
           "1. クリックポストで宛名を発行する" & vbCrLf & _
           "2. e飛伝で佐川伝票を印刷する" & vbCrLf & _
           "3. 追跡番号(伝票番号)を控えたら「実行3」を押してください。", _
           vbInformation, "実行2 完了"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "処理中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical, "エラー"
End Sub


'--------------------------------------------------
' 実行3: 出荷通知（※未確定。furusatoポータルへの通知方法を要確認）
'
'   ふるさと納税は出荷後にポータル側へ「伝票番号・配達業者」を登録する。
'   その入力フォーマット（CSV列・出力先）が現フォーマットからは確定できないため、
'   ここは骨組みのみ。実際の通知CSV仕様が分かり次第、他チャネルの手順7/8に倣って実装する。
'--------------------------------------------------
Sub 実行3_ふるさと納税_出荷通知()
    MsgBox "【未実装】出荷通知CSVの作成。" & vbCrLf & vbCrLf & _
           "ふるさと納税ポータルへの発送通知は、登録する列(伝票番号・配達業者など)と" & vbCrLf & _
           "出力フォーマットが確定してから実装します。" & vbCrLf & _
           "現在の通知のやり方（どの画面に何を入れているか）を教えてください。", _
           vbInformation, "実行3（準備中）"
End Sub


'==================================================
' 補助関数
'==================================================

'--- 謝礼品名から「佐川対象（重い/大型）」かどうかを判定 ---
'   下のキーワードに該当すれば佐川、それ以外はクリックポスト。
'   ※キーワードは要確認・要調整。重い商品や大型の謝礼品名に出てくる語を
'     カンマで足し引きするだけで判定を直せます。
Function Is佐川対象(giftName As String) As Boolean
    Dim keywords As Variant
    Dim k As Variant
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
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    If lr <= headerRow Then lr = headerRow + 1
    ws.Range(ws.Cells(headerRow + 1, 1), ws.Cells(ws.Rows.Count, lastCol)).ClearContents
End Sub
