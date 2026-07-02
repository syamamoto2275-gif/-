Attribute VB_Name = "ペライチ_マスター"
'==================================================
' ペライチ 出荷自動化 - マスター実行モジュール
'
' 【使い方】
'   1. ペライチからダウンロードしたCSVデータを元データシートに貼り付け
'   2. 実行1 → 発送方法をAR列に自動入力 + 編集シートに転記
'   3. 編集シートのA列（発送方法）を確認・修正
'   4. 実行2 → 各シート（物出し・アップデータ）にデータを作成
'   5. 追跡番号シートに追跡番号を貼り付けた後、実行3 → 通知用データシートに転記
'==================================================

Const COL_SHIPMETHOD As Integer = 44  ' AR列 = 発送方法
Const COL_ORDERNO As Integer = 3       ' C列  = 注文番号
Const COL_ITEMNAME As Integer = 8      ' H列  = 商品名
Const COL_CUST_LAST As Integer = 25    ' Y列  = 顧客名(姓)
Const COL_CUST_FIRST As Integer = 26   ' Z列  = 顧客名(名)
Const COL_CUST_POST As Integer = 29    ' AC列 = 郵便番号
Const COL_CUST_PREF As Integer = 30    ' AD列 = 都道府県
Const COL_CUST_ADDR1 As Integer = 31   ' AE列 = 住所1
Const COL_CUST_ADDR2 As Integer = 32   ' AF列 = 住所2
Const COL_CUST_TEL As Integer = 33     ' AG列 = 電話番号
Const COL_DEST_LAST As Integer = 35    ' AI列 = お届け先姓
Const COL_DEST_FIRST As Integer = 36   ' AJ列 = お届け先名
Const COL_DEST_POST As Integer = 37    ' AK列 = お届け先郵便番号
Const COL_DEST_PREF As Integer = 38    ' AL列 = お届け先都道府県
Const COL_DEST_ADDR1 As Integer = 39   ' AM列 = お届け先住所1
Const COL_DEST_ADDR2 As Integer = 40   ' AN列 = お届け先住所2
Const COL_DEST_TEL As Integer = 41     ' AO列 = お届け先電話番号


Function Is衣類ペライチ(itemName As String) As Boolean
    Dim keywords As Variant
    keywords = Array("脇高", "セクシー", "レース", "リボン", "ブラジャー", "ブラ", _
                     "インナー", "ショーツ", "下着", "ランジェリー", "パンティ", "ビキニ")
    Dim i As Integer
    For i = 0 To UBound(keywords)
        If InStr(itemName, keywords(i)) > 0 Then
            Is衣類ペライチ = True
            Exit Function
        End If
    Next i
    Is衣類ペライチ = False
End Function


Private Function Get元データSheet() As Worksheet
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If InStr(sh.Name, "元データ") > 0 Then
            Set Get元データSheet = sh
            Exit Function
        End If
    Next sh
    Set Get元データSheet = Nothing
End Function


Private Function Get配送情報(ws As Worksheet, rowNum As Long) As String()
    Dim result(5) As String

    Dim destLast As String
    destLast = Trim(ws.Cells(rowNum, COL_DEST_LAST).Value)

    If destLast <> "" Then
        result(0) = destLast & Trim(ws.Cells(rowNum, COL_DEST_FIRST).Value)
        result(1) = ws.Cells(rowNum, COL_DEST_POST).Value
        result(2) = ws.Cells(rowNum, COL_DEST_PREF).Value
        result(3) = ws.Cells(rowNum, COL_DEST_ADDR1).Value
        result(4) = ws.Cells(rowNum, COL_DEST_ADDR2).Value
        result(5) = ws.Cells(rowNum, COL_DEST_TEL).Value
    Else
        result(0) = Trim(ws.Cells(rowNum, COL_CUST_LAST).Value) & Trim(ws.Cells(rowNum, COL_CUST_FIRST).Value)
        result(1) = ws.Cells(rowNum, COL_CUST_POST).Value
        result(2) = ws.Cells(rowNum, COL_CUST_PREF).Value
        result(3) = ws.Cells(rowNum, COL_CUST_ADDR1).Value
        result(4) = ws.Cells(rowNum, COL_CUST_ADDR2).Value
        result(5) = ws.Cells(rowNum, COL_CUST_TEL).Value
    End If

    Get配送情報 = result
End Function


Private Sub シートクリア(ws As Worksheet, startRow As Long)
    If ws Is Nothing Then Exit Sub
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lr >= startRow Then ws.Rows(startRow & ":" & lr).ClearContents
End Sub


'--------------------------------------------------
' クリックポスト 一括発行アップデータの見出し（1行目）を正しい8列にそろえる
' クリックポストの「まとめ入力用CSV」は必ず次の8列。内容品は必ず8列目。
'   1 お届け先郵便番号 / 2 お届け先氏名 / 3 お届け先敬称
'   4 住所1行目 / 5 住所2行目 / 6 住所3行目 / 7 住所4行目(空でも可) / 8 内容品
' 見出しが7列だと「1行目の項目数が足りません」で弾かれるため、毎回ここで整える。
'--------------------------------------------------
Private Sub クリックPost見出し設定(ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    ws.Cells(1, 1).Value = "お届け先郵便番号"
    ws.Cells(1, 2).Value = "お届け先氏名"
    ws.Cells(1, 3).Value = "お届け先敬称"
    ws.Cells(1, 4).Value = "お届け先住所1行目"
    ws.Cells(1, 5).Value = "お届け先住所2行目"
    ws.Cells(1, 6).Value = "お届け先住所3行目"
    ws.Cells(1, 7).Value = "お届け先住所4行目"
    ws.Cells(1, 8).Value = "内容品"
End Sub


'--------------------------------------------------
' 佐川E飛伝アップデータをCSVファイルに出力する
'   保存先: …\Dropbox\ネットショップ\【出荷】ペライチ\e飛伝データ ペライチ\
'   ファイル名: ペライチ佐川YYYY.M.D.csv（Shift-JIS。社長の運用どおりの命名。月日は0埋めしない）
'   佐川データが1件以上（見出し行のみでない）ときだけ出力する。
'   列位置を崩さないため、間の空列も含めて1〜74列（e飛伝の正規フォーマット）をそのままCSV化する。
'   戻り値: 保存したフルパス（出力しなかった場合は空文字）
'--------------------------------------------------
Private Function 佐川CSV出力(wsSagawaUpd As Worksheet) As String
    佐川CSV出力 = ""
    If wsSagawaUpd Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = wsSagawaUpd.Cells(wsSagawaUpd.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function   ' 見出しだけ＝佐川データ無し

    Const LAST_COL As Long = 74         ' 編集１０まで（e飛伝の正規フォーマット74列）

    Dim folder As String
    folder = "C:\Users\" & Environ("USERNAME") & _
             "\Dropbox\ネットショップ\【出荷】ペライチ\e飛伝データ ペライチ\"
    If Dir(folder, vbDirectory) = "" Then
        MsgBox "佐川CSVの保存先フォルダが見つかりません：" & Chr(13) & folder & Chr(13) & Chr(13) & _
               "佐川シートは作成済みです。フォルダを確認してください。", vbExclamation
        Exit Function
    End If

    Dim filePath As String
    filePath = folder & "ペライチ佐川" & Year(Date) & "." & Month(Date) & "." & Day(Date) & ".csv"

    Dim fno As Integer: fno = FreeFile
    Open filePath For Output As #fno   ' 日本語Windowsでは Shift-JIS で書き出される
    Dim r As Long, c As Long, lineStr As String, cellStr As String
    For r = 1 To lastRow
        lineStr = ""
        For c = 1 To LAST_COL
            cellStr = CStr(wsSagawaUpd.Cells(r, c).Value)
            ' カンマ・改行・引用符を含む値はダブルクオートで囲む（CSVエスケープ）
            If InStr(cellStr, ",") > 0 Or InStr(cellStr, """") > 0 Or _
               InStr(cellStr, Chr(13)) > 0 Or InStr(cellStr, Chr(10)) > 0 Then
                cellStr = """" & Replace(cellStr, """", """""") & """"
            End If
            If c = 1 Then
                lineStr = cellStr
            Else
                lineStr = lineStr & "," & cellStr
            End If
        Next c
        Print #fno, lineStr
    Next r
    Close #fno

    佐川CSV出力 = filePath
End Function


'--------------------------------------------------
' 佐川E飛伝アップデータの見出し(1行目・74列)を正規フォーマットにそろえる
' （お手本 ペライチ佐川2026.6.29.csv と同一の並び）
'--------------------------------------------------
Private Sub 佐川見出し設定(ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    Dim h As Variant
    h = Array( _
        "お届け先コード取得区分", "お届け先コード", "お届け先電話番号", "お届け先郵便番号", _
        "お届け先住所１", "お届け先住所２", "お届け先住所３", "お届け先名称１", "お届け先名称２", _
        "お客様管理番号", "お客様コード", "部署ご担当者コード取得区分", "部署ご担当者コード", _
        "部署ご担当者名称", "荷送人電話番号", "ご依頼主コード取得区分", "ご依頼主コード", _
        "ご依頼主電話番号", "ご依頼主郵便番号", "ご依頼主住所１", "ご依頼主住所２", _
        "ご依頼主名称１", "ご依頼主名称２", "荷姿", "品名１", "品名２", "品名３", "品名４", "品名５", _
        "荷札荷姿", "荷札品名１", "荷札品名２", "荷札品名３", "荷札品名４", "荷札品名５", _
        "荷札品名６", "荷札品名７", "荷札品名８", "荷札品名９", "荷札品名１０", "荷札品名１１", _
        "出荷個数", "スピード指定", "クール便指定", "配達日", "配達指定時間帯", "配達指定時間（時分）", _
        "代引金額", "消費税", "決済種別", "保険金額", "指定シール１", "指定シール２", "指定シール３", _
        "営業所受取", "SRC区分", "営業所受取営業所コード", "元着区分", "メールアドレス", "ご不在時連絡先", _
        "出荷日", "お問い合せ送り状No.", "出荷場印字区分", "集約解除指定", _
        "編集０１", "編集０２", "編集０３", "編集０４", "編集０５", "編集０６", "編集０７", "編集０８", "編集０９", "編集１０")
    Dim i As Long
    For i = 0 To UBound(h)
        ws.Cells(1, i + 1).Value = h(i)
    Next i
End Sub


'--------------------------------------------------
' 佐川E飛伝アップデータへ1行書き込む（お手本6.29の正規フォーマット）
'   ご依頼主(shop fuu/小牧)・荷姿・出荷個数・スピード・クール便・指定シールは固定値。
'   お届け先・注文番号(お客様管理番号)・品名は引数で受ける。
'   コード類(001/000/013等)は先頭ゼロを保つため文字書式(@)で入れる。
'--------------------------------------------------
Private Sub 佐川行書込(ws As Worksheet, r As Long, orderNo As String, _
                       tel As String, post As String, pref As String, _
                       addr1 As String, addr2 As String, dName As String, hinmei As String)
    ' お届け先
    ws.Cells(r, 3).NumberFormat = "@": ws.Cells(r, 3).Value = tel        ' お届け先電話番号
    ws.Cells(r, 4).NumberFormat = "@": ws.Cells(r, 4).Value = post       ' お届け先郵便番号
    ws.Cells(r, 5).NumberFormat = "@": ws.Cells(r, 5).Value = pref       ' 住所１
    ws.Cells(r, 6).NumberFormat = "@": ws.Cells(r, 6).Value = addr1      ' 住所２
    ws.Cells(r, 7).NumberFormat = "@": ws.Cells(r, 7).Value = addr2      ' 住所３
    ws.Cells(r, 8).Value = dName                                          ' 名称１
    ws.Cells(r, 10).NumberFormat = "@": ws.Cells(r, 10).Value = orderNo  ' お客様管理番号＝注文番号
    ' ご依頼主（固定：shop fuu／株式会社SEED 小牧）
    ws.Cells(r, 18).NumberFormat = "@": ws.Cells(r, 18).Value = "0568-47-5090"
    ws.Cells(r, 19).NumberFormat = "@": ws.Cells(r, 19).Value = "485-0814"
    ws.Cells(r, 20).Value = "愛知県小牧市"
    ws.Cells(r, 21).Value = "古雅３－６１－４"
    ws.Cells(r, 22).Value = "ｓｈｏｐ　ｆｕｕ"
    ' 荷姿・品名・出荷個数・スピード・クール便・指定シール（固定：お手本6.29に一致）
    ws.Cells(r, 24).NumberFormat = "@": ws.Cells(r, 24).Value = "001"    ' 荷姿
    ws.Cells(r, 25).Value = hinmei                                        ' 品名１
    ws.Cells(r, 42).Value = 1                                             ' 出荷個数
    ws.Cells(r, 43).NumberFormat = "@": ws.Cells(r, 43).Value = "000"    ' スピード指定
    ws.Cells(r, 44).NumberFormat = "@": ws.Cells(r, 44).Value = "001"    ' クール便指定（お手本6.29に合わせる）
    ws.Cells(r, 52).NumberFormat = "@": ws.Cells(r, 52).Value = "013"    ' 指定シール１
    ws.Cells(r, 53).NumberFormat = "@": ws.Cells(r, 53).Value = "011"    ' 指定シール２
End Sub


'--------------------------------------------------
' 実行1: 発送方法をAR列に自動入力 + 編集シートへ転記
'--------------------------------------------------
Sub 実行1_ペライチ_発送方法自動判定()
    Dim ws As Worksheet
    Set ws = Get元データSheet()

    If ws Is Nothing Then
        MsgBox "元データシートが見つかりません。", vbCritical, "エラー"
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_ORDERNO).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "データがありません。先にペライチの注文データを貼り付けてください。", vbExclamation
        Exit Sub
    End If

    ws.Cells(1, COL_SHIPMETHOD).Value = "発送方法"

    Application.ScreenUpdating = False

    Dim i As Long
    For i = 2 To lastRow
        If Trim(ws.Cells(i, COL_ORDERNO).Value) = "" Then GoTo Skip1

        Dim itemName As String
        itemName = ws.Cells(i, COL_ITEMNAME).Value

        Dim destName As String, destPostal As String
        If Trim(ws.Cells(i, COL_DEST_LAST).Value) <> "" Then
            destName = Trim(ws.Cells(i, COL_DEST_LAST).Value) & Trim(ws.Cells(i, COL_DEST_FIRST).Value)
            destPostal = ws.Cells(i, COL_DEST_POST).Value
        Else
            destName = Trim(ws.Cells(i, COL_CUST_LAST).Value) & Trim(ws.Cells(i, COL_CUST_FIRST).Value)
            destPostal = ws.Cells(i, COL_CUST_POST).Value
        End If

        Dim cnt As Long
        cnt = 0
        Dim j As Long
        For j = 2 To lastRow
            Dim jName As String, jPostal As String
            If Trim(ws.Cells(j, COL_DEST_LAST).Value) <> "" Then
                jName = Trim(ws.Cells(j, COL_DEST_LAST).Value) & Trim(ws.Cells(j, COL_DEST_FIRST).Value)
                jPostal = ws.Cells(j, COL_DEST_POST).Value
            Else
                jName = Trim(ws.Cells(j, COL_CUST_LAST).Value) & Trim(ws.Cells(j, COL_CUST_FIRST).Value)
                jPostal = ws.Cells(j, COL_CUST_POST).Value
            End If
            If jName = destName And jPostal = destPostal Then cnt = cnt + 1
        Next j

        If Is衣類ペライチ(itemName) Then
            If cnt >= 7 Then
                ws.Cells(i, COL_SHIPMETHOD).Value = "衣類佐川"
            ElseIf cnt > 3 Then
                ws.Cells(i, COL_SHIPMETHOD).Value = "衣類クリック複数"
            Else
                ws.Cells(i, COL_SHIPMETHOD).Value = "衣類クリック1通"
            End If
        Else
            Dim isPetClick As Boolean
            isPetClick = False
            If InStr(itemName, "ごはん") > 0 Or InStr(itemName, "おやつ") > 0 Then
                If InStr(itemName, "3kg") = 0 And InStr(itemName, "5kg") = 0 And InStr(itemName, "10kg") = 0 And _
                   InStr(itemName, "3ｋｇ") = 0 And InStr(itemName, "5ｋｇ") = 0 And InStr(itemName, "10ｋｇ") = 0 Then
                    isPetClick = True
                End If
            End If
            If isPetClick Then
                ws.Cells(i, COL_SHIPMETHOD).Value = IIf(cnt > 1, "ペットクリック複数", "ペットクリック1通")
            Else
                ws.Cells(i, COL_SHIPMETHOD).Value = "ペット佐川"
            End If
        End If

        Skip1:
    Next i

    Call ペライチ_編集シート転記(ws, lastRow)

    Application.ScreenUpdating = True

    ' 実行1の後は編集シートを表示する（A列の発送方法を確認してもらうため）(社長指示 2026-07-01)
    Dim shEdit As Worksheet
    For Each shEdit In ThisWorkbook.Sheets
        If shEdit.Name = "編集" Then shEdit.Activate: Exit For
    Next shEdit

    MsgBox "発送方法をAR列に自動入力し、編集シートにデータを転記しました。" & Chr(13) & Chr(13) & _
           "【確認してください】" & Chr(13) & _
           "・編集シートのA列で発送方法を確認" & Chr(13) & _
           "・佐川急便に変更したい行は「ペット佐川」または「衣類佐川」に変更" & Chr(13) & Chr(13) & _
           "確認後「実行2」ボタンを押してください。", _
           vbInformation, "実行1 完了"
End Sub


'--------------------------------------------------
' 元データ → 編集シートへデータ転記
'--------------------------------------------------
Sub ペライチ_編集シート転記(ws As Worksheet, lastRow As Long)
    Dim wsEdit As Worksheet
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = "編集" Then Set wsEdit = sh
    Next sh
    If wsEdit Is Nothing Then Exit Sub

    Call シートクリア(wsEdit, 2)

    Dim editRow As Long
    editRow = 2

    Dim i As Long
    For i = 2 To lastRow
        If Trim(ws.Cells(i, COL_ORDERNO).Value) = "" Then GoTo SkipEdit

        Dim info() As String
        info = Get配送情報(ws, i)

        ' 郵便・住所・電話は「文字」書式にしてから入れる。
        ' これをしないと「443」等の数字だけの住所をExcelが日付(1901/3/18)に化けさせる(社長指摘 2026-07-01)
        wsEdit.Cells(editRow, "E").NumberFormat = "@"  ' 郵便番号
        wsEdit.Cells(editRow, "G").NumberFormat = "@"  ' 住所1
        wsEdit.Cells(editRow, "H").NumberFormat = "@"  ' 住所2
        wsEdit.Cells(editRow, "I").NumberFormat = "@"  ' 電話番号

        wsEdit.Cells(editRow, "A").Value = ws.Cells(i, COL_SHIPMETHOD).Value
        wsEdit.Cells(editRow, "B").Value = ws.Cells(i, COL_ORDERNO).Value
        wsEdit.Cells(editRow, "C").Value = ws.Cells(i, COL_ITEMNAME).Value
        wsEdit.Cells(editRow, "D").Value = info(0)
        wsEdit.Cells(editRow, "E").Value = info(1)
        wsEdit.Cells(editRow, "F").Value = info(2)
        wsEdit.Cells(editRow, "G").Value = info(3)
        wsEdit.Cells(editRow, "H").Value = info(4)
        wsEdit.Cells(editRow, "I").Value = info(5)
        wsEdit.Cells(editRow, "J").Value = ws.Cells(i, 9).Value
        wsEdit.Cells(editRow, "K").Value = ws.Cells(i, 42).Value

        editRow = editRow + 1
        SkipEdit:
    Next i
End Sub


'--------------------------------------------------
' 実行2: 各シートにデータを作成
'   編集シートのデータを発送方法に応じて振り分ける
'--------------------------------------------------
Sub 実行2_ペライチ_出荷CSV作成()
    ' 編集シートを取得
    Dim wsEdit As Worksheet
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = "編集" Then Set wsEdit = sh
    Next sh

    If wsEdit Is Nothing Then
        MsgBox "編集シートが見つかりません。先に実行1を押してください。", vbCritical, "エラー"
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = wsEdit.Cells(wsEdit.Rows.Count, "B").End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "編集シートにデータがありません。先に実行1を押してください。", vbExclamation
        Exit Sub
    End If

    ' 各シートを取得
    Dim wsPetClick As Worksheet
    Dim wsPetSagawa As Worksheet
    Dim wsIrui As Worksheet
    Dim wsPetClickUpd As Worksheet
    Dim wsIruiClickUpd As Worksheet
    Dim wsSagawaUpd As Worksheet

    For Each sh In ThisWorkbook.Sheets
        Dim n As String
        n = sh.Name
        If n = "ペットクリック物出し" Then Set wsPetClick = sh
        If n = "ペット佐川物出し" Then Set wsPetSagawa = sh
        If n = "衣類物出し" Then Set wsIrui = sh
        If InStr(n, "ペット_クリック") > 0 And InStr(n, "一括発行アップデータ") > 0 Then Set wsPetClickUpd = sh
        If InStr(n, "衣類_クリック") > 0 And InStr(n, "一括発行アップデータ") > 0 Then Set wsIruiClickUpd = sh
        If InStr(n, "佐川発行用") > 0 And InStr(n, "E飛伝") > 0 Then Set wsSagawaUpd = sh
    Next sh

    Application.ScreenUpdating = False

    ' 各シートのデータをクリア（ヘッダー行を残す）
    Call シートクリア(wsPetClick, 2)
    Call シートクリア(wsPetSagawa, 2)
    Call シートクリア(wsIrui, 2)
    Call シートクリア(wsPetClickUpd, 2)
    Call シートクリア(wsIruiClickUpd, 2)

    ' クリックポスト一括発行アップデータの見出しを正しい8列に整える（内容品を8列目に）
    Call クリックPost見出し設定(wsPetClickUpd)
    Call クリックPost見出し設定(wsIruiClickUpd)

    ' 佐川シートを作り直す：行1=見出し(74列 正規フォーマット)・行2以降=データ (社長指示 2026-07-02)
    If Not wsSagawaUpd Is Nothing Then
        Dim lrSg As Long
        lrSg = wsSagawaUpd.Cells(wsSagawaUpd.Rows.Count, 1).End(xlUp).Row
        If lrSg >= 1 Then wsSagawaUpd.Rows("1:" & lrSg).ClearContents
        Call 佐川見出し設定(wsSagawaUpd)
    End If

    ' 書き込み行カウンター
    Dim rPC As Long: rPC = 2
    Dim rPS As Long: rPS = 2
    Dim rIR As Long: rIR = 2
    Dim rPCU As Long: rPCU = 2
    Dim rICU As Long: rICU = 2
    Dim rSG As Long: rSG = 2   ' 佐川データは行2から(#4)

    Dim c1 As Long, c2 As Long, c3 As Long, c4 As Long

    Dim i As Long
    For i = 2 To lastRow
        Dim method As String
        method = Trim(wsEdit.Cells(i, "A").Value)
        If method = "" Then GoTo Skip2

        Dim orderNo As String: orderNo = wsEdit.Cells(i, "B").Value
        Dim iName As String: iName = wsEdit.Cells(i, "C").Value
        Dim dName As String: dName = wsEdit.Cells(i, "D").Value
        Dim post As String: post = wsEdit.Cells(i, "E").Value
        Dim pref As String: pref = wsEdit.Cells(i, "F").Value
        Dim addr1 As String: addr1 = wsEdit.Cells(i, "G").Value
        Dim addr2 As String: addr2 = wsEdit.Cells(i, "H").Value
        Dim tel As String: tel = wsEdit.Cells(i, "I").Value

        ' 内容品・品名はクリック/佐川とも「shop fuu / ペット用品」に統一
        ' （社長指示 2026-07-02：ペライチは衣類の販売なし）
        Dim prod As String
        prod = "shop fuu / ペット用品"

        If InStr(method, "ペット") > 0 And InStr(method, "クリック") > 0 Then
            ' ペットクリック物出し
            If Not wsPetClick Is Nothing Then
                wsPetClick.Cells(rPC, 5).NumberFormat = "@"  ' 郵便
                wsPetClick.Cells(rPC, 7).NumberFormat = "@"  ' 住所2(数字だけでも日付化しない)
                wsPetClick.Cells(rPC, 8).NumberFormat = "@"  ' 電話
                wsPetClick.Cells(rPC, 1).Value = method
                wsPetClick.Cells(rPC, 2).Value = orderNo
                wsPetClick.Cells(rPC, 3).Value = iName
                wsPetClick.Cells(rPC, 4).Value = dName
                wsPetClick.Cells(rPC, 5).Value = post
                wsPetClick.Cells(rPC, 6).Value = pref & addr1
                wsPetClick.Cells(rPC, 7).Value = addr2
                wsPetClick.Cells(rPC, 8).Value = tel
                rPC = rPC + 1
            End If
            ' ペット_クリック一括発行アップデータ（クリックポスト8列：内容品は8列目）
            If Not wsPetClickUpd Is Nothing Then
                wsPetClickUpd.Cells(rPCU, 1).NumberFormat = "@"  ' 郵便
                wsPetClickUpd.Cells(rPCU, 6).NumberFormat = "@"  ' 住所3行目(数字だけでも日付化しない)
                wsPetClickUpd.Cells(rPCU, 1).Value = post
                wsPetClickUpd.Cells(rPCU, 2).Value = dName
                wsPetClickUpd.Cells(rPCU, 3).Value = "様"
                wsPetClickUpd.Cells(rPCU, 4).Value = pref     ' 住所1行目
                wsPetClickUpd.Cells(rPCU, 5).Value = addr1    ' 住所2行目
                wsPetClickUpd.Cells(rPCU, 6).Value = addr2    ' 住所3行目
                wsPetClickUpd.Cells(rPCU, 7).Value = ""       ' 住所4行目(未使用でも8列必須なので空で置く)
                wsPetClickUpd.Cells(rPCU, 8).Value = prod     ' 内容品(必ず8列目)
                rPCU = rPCU + 1
            End If
            c1 = c1 + 1

        ElseIf InStr(method, "衣類") > 0 And InStr(method, "クリック") > 0 Then
            ' 衣類物出し
            If Not wsIrui Is Nothing Then
                wsIrui.Cells(rIR, 1).Value = method
                wsIrui.Cells(rIR, 2).Value = orderNo
                wsIrui.Cells(rIR, 3).Value = iName
                wsIrui.Cells(rIR, 4).Value = dName
                wsIrui.Cells(rIR, 5).Value = post
                wsIrui.Cells(rIR, 6).Value = pref & addr1
                wsIrui.Cells(rIR, 7).Value = addr2
                wsIrui.Cells(rIR, 8).Value = tel
                rIR = rIR + 1
            End If
            ' 衣類_クリック一括発行アップデータ（クリックポスト8列：内容品は8列目）
            If Not wsIruiClickUpd Is Nothing Then
                wsIruiClickUpd.Cells(rICU, 1).NumberFormat = "@"  ' 郵便
                wsIruiClickUpd.Cells(rICU, 6).NumberFormat = "@"  ' 住所3行目(数字だけでも日付化しない)
                wsIruiClickUpd.Cells(rICU, 1).Value = post
                wsIruiClickUpd.Cells(rICU, 2).Value = dName
                wsIruiClickUpd.Cells(rICU, 3).Value = "様"
                wsIruiClickUpd.Cells(rICU, 4).Value = pref     ' 住所1行目
                wsIruiClickUpd.Cells(rICU, 5).Value = addr1    ' 住所2行目
                wsIruiClickUpd.Cells(rICU, 6).Value = addr2    ' 住所3行目
                wsIruiClickUpd.Cells(rICU, 7).Value = ""       ' 住所4行目(未使用でも8列必須なので空で置く)
                wsIruiClickUpd.Cells(rICU, 8).Value = prod     ' 内容品(必ず8列目)
                rICU = rICU + 1
            End If
            c2 = c2 + 1

        ElseIf method = "ペット佐川" Then
            ' ペット佐川物出し
            If Not wsPetSagawa Is Nothing Then
                wsPetSagawa.Cells(rPS, 1).Value = method
                wsPetSagawa.Cells(rPS, 2).Value = orderNo
                wsPetSagawa.Cells(rPS, 3).Value = iName
                wsPetSagawa.Cells(rPS, 4).Value = dName
                wsPetSagawa.Cells(rPS, 5).Value = post
                wsPetSagawa.Cells(rPS, 6).Value = pref & addr1
                wsPetSagawa.Cells(rPS, 7).Value = addr2
                wsPetSagawa.Cells(rPS, 8).Value = tel
                rPS = rPS + 1
            End If
            ' 佐川E飛伝アップデータ（正規フォーマットで1行書込）
            If Not wsSagawaUpd Is Nothing Then
                Call 佐川行書込(wsSagawaUpd, rSG, orderNo, tel, post, pref, addr1, addr2, dName, prod)
                rSG = rSG + 1
            End If
            c3 = c3 + 1

        ElseIf method = "衣類佐川" Then
            ' 衣類物出し
            If Not wsIrui Is Nothing Then
                wsIrui.Cells(rIR, 1).Value = method
                wsIrui.Cells(rIR, 2).Value = orderNo
                wsIrui.Cells(rIR, 3).Value = iName
                wsIrui.Cells(rIR, 4).Value = dName
                wsIrui.Cells(rIR, 5).Value = post
                wsIrui.Cells(rIR, 6).Value = pref & addr1
                wsIrui.Cells(rIR, 7).Value = addr2
                wsIrui.Cells(rIR, 8).Value = tel
                rIR = rIR + 1
            End If
            ' 佐川E飛伝アップデータ（正規フォーマットで1行書込）※ペライチは衣類販売なしのため通常ここは通らない
            If Not wsSagawaUpd Is Nothing Then
                Call 佐川行書込(wsSagawaUpd, rSG, orderNo, tel, post, pref, addr1, addr2, dName, prod)
                rSG = rSG + 1
            End If
            c4 = c4 + 1
        End If

        Skip2:
    Next i

    Application.ScreenUpdating = True

    Dim msg As String
    msg = "各シートへのデータ作成が完了しました！" & Chr(13) & Chr(13)
    If c1 > 0 Then msg = msg & "ペットクリックポスト: " & c1 & "件" & Chr(13)
    If c2 > 0 Then msg = msg & "衣類クリックポスト: " & c2 & "件" & Chr(13)
    If c3 > 0 Then msg = msg & "ペット佐川: " & c3 & "件" & Chr(13)
    If c4 > 0 Then msg = msg & "衣類佐川: " & c4 & "件" & Chr(13)
    If c1 + c2 + c3 + c4 = 0 Then msg = msg & "出力できるデータがありませんでした。" & Chr(13) & "編集シートのA列（発送方法）を確認してください。"
    msg = msg & Chr(13) & "各シートを確認してください。"

    ' 納品書シートにデータを作成
    Call ペライチ_納品書作成(wsEdit, lastRow)

    ' 佐川E飛伝アップデータをCSVファイルに出力（佐川が1件以上のときだけ）(社長指示 2026-07-02)
    If (c3 + c4) > 0 Then
        Dim sgCsvPath As String
        sgCsvPath = 佐川CSV出力(wsSagawaUpd)
        If sgCsvPath <> "" Then
            msg = msg & Chr(13) & "佐川E飛伝CSVを保存しました：" & Chr(13) & sgCsvPath & Chr(13)
        End If
    End If

    ' 実行2の後はペットクリック物出しシートを表示する(#3 社長指示 2026-07-01)
    If Not wsPetClick Is Nothing Then wsPetClick.Activate

    MsgBox msg, vbInformation, "実行2 完了"
End Sub


'--------------------------------------------------
' 納品書シートへのデータ書き込み
' 編集シートの行2 → 納品書【2行目】、行3 → 納品書【3行目】...
'--------------------------------------------------
Sub ペライチ_納品書作成(wsEdit As Worksheet, lastRow As Long)
    Dim wsGen As Worksheet
    Set wsGen = Get元データSheet()

    ' 全納品書シートをクリア
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name Like "納品書【*行目】" Then
            sh.Cells(5, 7).Value = ""   ' 注文番号
            sh.Cells(5, 10).Value = ""  ' 注文日
            ' 顧客側はB列に集約（C列は使わない＝空にする）
            sh.Cells(11, 2).Value = ""  ' 顧客郵便番号(B11)
            sh.Cells(11, 3).ClearContents  ' C11は使わない(空文字だとB11がはみ出せず切れるのでClearContents)
            sh.Cells(12, 2).Value = ""  ' 顧客住所1
            sh.Cells(13, 2).Value = ""  ' 顧客住所2
            sh.Cells(14, 2).Value = "様"  ' 顧客名リセット(B14)
            sh.Cells(16, 2).Value = ""  ' 顧客TEL(B16)
            sh.Cells(16, 3).ClearContents  ' C16は使わない(空文字だとB16のTELが切れるのでClearContents)
            sh.Cells(17, 2).Value = ""  ' 顧客Mail(B17)
            sh.Cells(17, 3).ClearContents  ' C17は使わない(空文字だとB17のMailが切れるのでClearContents)
            ' お届け先側はG列に集約（H列は使わない＝空にする）
            sh.Cells(11, 7).Value = ""  ' お届け先郵便番号(G11)
            sh.Cells(11, 8).ClearContents  ' H11は使わない(空文字だとG11がはみ出せず切れるのでClearContents)
            sh.Cells(12, 7).Value = ""  ' お届け先住所1
            sh.Cells(13, 7).Value = ""  ' お届け先住所2
            sh.Cells(14, 7).Value = "様"  ' お届け先名リセット(G14)
            sh.Cells(16, 7).Value = ""  ' お届け先TEL(G16)
            sh.Cells(16, 8).ClearContents  ' H16は使わない(空文字だとG16のTELが切れるのでClearContents)
            sh.Cells(17, 7).Value = ""  ' お届け先Mail(G17)リセット(#7)
            sh.Cells(19, 4).Value = ""  ' 総合計金額
            sh.Cells(22, 8).Value = ""  ' 単価(H22)…実行毎に入れ直すのでリセット(#1 社長指示 2026-07-01)
            Dim r As Long
            For r = 22 To 31
                sh.Cells(r, 2).Value = ""   ' 商品名
                sh.Cells(r, 9).Value = ""   ' 数量
                sh.Cells(r, 10).Value = ""  ' 小計
            Next r
            sh.Cells(34, 2).Value = ""  ' 備考
            sh.Cells(34, 10).Value = "" ' 送料
            sh.Cells(37, 9).Value = ""  ' お支払い方法(I37)…外部リンクを使わず実データで入れ直す(#8)
        End If
    Next sh

    Dim i As Long
    For i = 2 To lastRow
        Dim orderNo As String
        orderNo = wsEdit.Cells(i, "B").Value
        If orderNo = "" Then GoTo NextNob

        ' 対応する納品書シートを取得
        Dim nobName As String
        nobName = "納品書【" & i & "行目】"
        Dim wsNob As Worksheet
        Set wsNob = Nothing
        For Each sh In ThisWorkbook.Sheets
            If sh.Name = nobName Then Set wsNob = sh: Exit For
        Next sh
        If wsNob Is Nothing Then GoTo NextNob

        ' 元データから顧客情報を取得（注文番号でマッチング）
        Dim genRow As Long
        genRow = 0
        If Not wsGen Is Nothing Then
            Dim k As Long
            For k = 2 To wsGen.Cells(wsGen.Rows.Count, COL_ORDERNO).End(xlUp).Row
                If CStr(wsGen.Cells(k, COL_ORDERNO).Value) = CStr(orderNo) Then
                    genRow = k
                    Exit For
                End If
            Next k
        End If

        ' 発行日・注文番号・注文日
        wsNob.Cells(3, 10).Value = Date
        wsNob.Cells(3, 10).NumberFormat = "yyyy/mm/dd"
        wsNob.Cells(5, 7).Value = orderNo

        If genRow > 0 Then
            wsNob.Cells(5, 10).Value = wsGen.Cells(genRow, 6).Value  ' 注文日時

            ' 顧客情報
            Dim custName As String
            custName = Trim(wsGen.Cells(genRow, COL_CUST_LAST).Value) & _
                       Trim(wsGen.Cells(genRow, COL_CUST_FIRST).Value)
            ' 顧客：郵便番号は先頭に「〒」、名前は末尾に「様」、TEL/Mailは接頭辞付き（すべてB列）
            wsNob.Cells(11, 2).NumberFormat = "@"
            wsNob.Cells(11, 2).Value = "〒" & CStr(wsGen.Cells(genRow, COL_CUST_POST).Value)
            wsNob.Cells(12, 2).NumberFormat = "@"  ' 住所は数字だけでも文字扱い(日付化防止)
            wsNob.Cells(13, 2).NumberFormat = "@"
            wsNob.Cells(12, 2).Value = wsGen.Cells(genRow, COL_CUST_PREF).Value & _
                                        wsGen.Cells(genRow, COL_CUST_ADDR1).Value
            wsNob.Cells(13, 2).Value = wsGen.Cells(genRow, COL_CUST_ADDR2).Value
            wsNob.Cells(14, 2).Value = custName & "様"
            wsNob.Cells(16, 2).NumberFormat = "@"
            wsNob.Cells(16, 2).Value = "TEL：" & CStr(wsGen.Cells(genRow, COL_CUST_TEL).Value)
            wsNob.Cells(17, 2).NumberFormat = "@"
            wsNob.Cells(17, 2).Value = "Mail：" & CStr(wsGen.Cells(genRow, 34).Value)  ' AH=メール
            wsNob.Cells(34, 10).Value = wsGen.Cells(genRow, 10).Value  ' J=送料
            ' お支払い方法(I37)…元データV列=支払い方法を反映（外部リンクの代わり）(#8)
            wsNob.Cells(37, 9).Value = wsGen.Cells(genRow, 22).Value
            ' ※お届け先G17のメールは表示しない（社長指示 2026-07-01：G17メール不要）
        End If

        ' お届け先情報（編集シートから）：郵便番号「〒」、名前「様」、TEL接頭辞付き（すべてG列）
        wsNob.Cells(11, 7).NumberFormat = "@"
        wsNob.Cells(11, 7).Value = "〒" & CStr(wsEdit.Cells(i, "E").Value)
        wsNob.Cells(12, 7).NumberFormat = "@"  ' 住所は数字だけでも文字扱い(日付化防止)
        wsNob.Cells(13, 7).NumberFormat = "@"
        wsNob.Cells(12, 7).Value = wsEdit.Cells(i, "F").Value & wsEdit.Cells(i, "G").Value
        wsNob.Cells(13, 7).Value = wsEdit.Cells(i, "H").Value
        wsNob.Cells(14, 7).Value = wsEdit.Cells(i, "D").Value & "様"
        wsNob.Cells(16, 7).NumberFormat = "@"
        wsNob.Cells(16, 7).Value = "TEL：" & CStr(wsEdit.Cells(i, "I").Value)

        ' 商品情報（1行目）
        Dim itemName As String: itemName = wsEdit.Cells(i, "C").Value
        Dim itemAmt As Variant: itemAmt = wsEdit.Cells(i, "J").Value
        If IsEmpty(itemAmt) Or Not IsNumeric(itemAmt) Then itemAmt = 0
        wsNob.Cells(22, 2).Value = itemName
        wsNob.Cells(22, 2).WrapText = False   ' 折り返し解除
        wsNob.Cells(22, 8).Value = CLng(itemAmt)  ' 単価
        wsNob.Cells(22, 9).Value = 1               ' 数量
        wsNob.Cells(22, 10).Value = CLng(itemAmt)  ' 小計

        ' 総合計金額
        Dim soryo As Variant: soryo = wsNob.Cells(34, 10).Value
        If IsEmpty(soryo) Or Not IsNumeric(soryo) Then soryo = 0
        wsNob.Cells(19, 4).Value = CLng(itemAmt) + CLng(soryo)

        NextNob:
    Next i
End Sub


'--------------------------------------------------
' 実行3: 追跡番号シートから通知用データシートに転記
'--------------------------------------------------
Sub 実行3_ペライチ_配送情報CSV作成()
    Dim wsTrack As Worksheet
    Dim wsNotify As Worksheet
    Dim sh As Worksheet

    For Each sh In ThisWorkbook.Sheets
        If InStr(sh.Name, "追跡番号") > 0 And InStr(sh.Name, "レポート") = 0 Then Set wsTrack = sh
        If InStr(sh.Name, "通知用データ") > 0 Then Set wsNotify = sh
    Next sh

    If wsTrack Is Nothing Then
        MsgBox "追跡番号シートが見つかりません。", vbCritical, "エラー"
        Exit Sub
    End If

    If wsNotify Is Nothing Then
        MsgBox "通知用データシートが見つかりません。", vbCritical, "エラー"
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = wsTrack.Cells(wsTrack.Rows.Count, "A").End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "追跡番号シートにデータがありません。" & Chr(13) & _
               "追跡番号シートに以下の順番でデータを入力してください：" & Chr(13) & _
               "  A列: 注文番号" & Chr(13) & _
               "  B列: 運送会社" & Chr(13) & _
               "  C列: 送り状番号（追跡番号）" & Chr(13) & _
               "  D列: 到着予定日" & Chr(13) & _
               "  E列: 到着時間From" & Chr(13) & _
               "  F列: 到着時間To", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Call シートクリア(wsNotify, 2)

    Dim cnt As Long
    cnt = 0
    Dim i As Long
    For i = 2 To lastRow
        If Trim(wsTrack.Cells(i, "A").Value) = "" Then GoTo Skip3

        wsNotify.Cells(cnt + 2, 1).Value = wsTrack.Cells(i, "A").Value
        wsNotify.Cells(cnt + 2, 2).Value = wsTrack.Cells(i, "B").Value
        wsNotify.Cells(cnt + 2, 3).Value = wsTrack.Cells(i, "C").Value
        wsNotify.Cells(cnt + 2, 4).Value = wsTrack.Cells(i, "D").Text
        wsNotify.Cells(cnt + 2, 5).Value = wsTrack.Cells(i, "E").Text
        wsNotify.Cells(cnt + 2, 6).Value = wsTrack.Cells(i, "F").Text
        cnt = cnt + 1
        Skip3:
    Next i

    Application.ScreenUpdating = True

    MsgBox "通知用データシートにデータを転記しました。" & Chr(13) & Chr(13) & _
           "件数: " & cnt & "件" & Chr(13) & Chr(13) & _
           "「通知用データ」シートを確認してください。", _
           vbInformation, "実行3 完了"
End Sub
