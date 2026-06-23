Attribute VB_Name = "Amazon_マスター"
'==================================================
' Amazon出荷自動化 - マスター実行モジュール
' ボタン3つで全工程を制御する
'
' 【ボタン配置】
'   実行1_Amazon_編集データ作成   → Sub 実行1_Amazon_編集データ作成
'   実行2_Amazon_出荷データ作成   → Sub 実行2_Amazon_出荷データ作成
'   実行3_Amazon_通知CSV作成      → Sub 実行3_Amazon_通知CSV作成
'==================================================


'--------------------------------------------------
' 実行1: 元データ → 編集シート作成 + 発送方法自動判定
'--------------------------------------------------
Sub 実行1_Amazon_編集データ作成()
    Application.ScreenUpdating = False

    ' 既存マクロ: 元データ → 編集シートへコピー
    Call 手順1_元データから編集データ作成

    ' 発送方法をA列に自動入力
    Call Amazon_発送方法自動判定

    Application.ScreenUpdating = True

    MsgBox "編集データを作成しました。" & Chr(13) & Chr(13) & _
           "【次のステップ】" & Chr(13) & _
           "「編集」シートのA列（発送方法）を確認・修正してください。" & Chr(13) & _
           "コスト面で佐川の方が安い場合は「ペット佐川」「衣類佐川」に変更できます。" & Chr(13) & Chr(13) & _
           "確認が完了したら「実行2」ボタンを押してください。", _
           vbInformation, "実行1 完了"
End Sub


'--------------------------------------------------
' 発送方法の自動判定（編集シートA列に書き込み）
'
' 判定ルール:
'   商品名に衣類キーワードが含まれる → 衣類クリック
'   それ以外 → ペットクリック
'   同一送付先（氏名+住所）が複数件 → 複数
'   それ以外 → 1通
'   佐川への変更は手動で行う
'--------------------------------------------------
Sub Amazon_発送方法自動判定()
    Dim ws As Worksheet
    Set ws = Sheets("編集")

    ' C列(order-id)で最終行を取得
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    If lastRow < 2 Then Exit Sub

    Dim i As Long
    For i = 2 To lastRow
        Dim productName As String
        productName = ws.Cells(i, "N").Value  ' N列 = product-name

        Dim recipientName As String
        recipientName = ws.Cells(i, "S").Value  ' S列 = recipient-name

        Dim address1 As String
        address1 = ws.Cells(i, "T").Value  ' T列 = ship-address-1

        ' 同一送付先の件数（氏名 + 住所1 で判断）
        Dim cnt As Long
        cnt = Application.WorksheetFunction.CountIfs( _
            ws.Columns("S"), recipientName, _
            ws.Columns("T"), address1)

        ' 衣類/ペット判定
        Dim isClothing As Boolean
        isClothing = Is衣類Amazon(productName)

        ' A列に発送方法をセット
        If isClothing Then
            ws.Cells(i, "A").Value = IIf(cnt > 1, "衣類クリック複数", "衣類クリック1通")
        Else
            ws.Cells(i, "A").Value = IIf(cnt > 1, "ペットクリック複数", "ペットクリック1通")
        End If
    Next i
End Sub


'--------------------------------------------------
' 実行2: 出荷データ作成一式（手順2〜5）
'   - 編集シートから各「編集から貼付」シートへ自動振り分け
'   - 物出しリスト作成
'   - クリックポスト・佐川アップデータ作成
'--------------------------------------------------
Sub 実行2_Amazon_出荷データ作成()
    Application.ScreenUpdating = False

    ' 編集シートから各発行用シートへ自動振り分け
    Call Amazon_振分_編集から貼付

    ' 既存マクロを順番に実行
    Call 手順2_物出しリスト作成
    Call 手順3_ペット_クリック発行用_一括発行アップデータ
    Call 手順4_衣類_クリック発行用_一括発行アップデータ
    Call 手順5_佐川発行用_E飛伝アップデータ

    Application.ScreenUpdating = True

    MsgBox "出荷データを作成しました。" & Chr(13) & Chr(13) & _
           "【次のステップ】" & Chr(13) & _
           "1. クリックポストで宛名を発行する" & Chr(13) & _
           "2. e飛伝で佐川伝票を印刷する" & Chr(13) & _
           "3. 追跡番号レポートを各シートに貼り付ける" & Chr(13) & Chr(13) & _
           "完了したら「実行3」ボタンを押してください。", _
           vbInformation, "実行2 完了"
End Sub


'--------------------------------------------------
' 編集シートの各行を発送方法に応じて振り分ける
'
' A列の値 → コピー先シート:
'   ペットクリック1通 / ペットクリック複数 → ペット_クリック発行用_編集から貼付
'   衣類クリック1通 / 衣類クリック複数     → 衣類_クリック発行用_編集から貼付
'   ペット佐川 / 衣類佐川                  → 佐川発行用_編集から貼付
'--------------------------------------------------
Sub Amazon_振分_編集から貼付()
    Dim wsEdit As Worksheet
    Set wsEdit = Sheets("編集")

    Dim wsPetClick As Worksheet
    Dim wsIruiClick As Worksheet
    Dim wsSagawa As Worksheet
    Set wsPetClick = Sheets("ペット_クリック発行用_編集から貼付")
    Set wsIruiClick = Sheets("衣類_クリック発行用_編集から貼付")
    Set wsSagawa = Sheets("佐川発行用_編集から貼付")

    ' 各シートの既存データをクリア（ヘッダー行は残す）
    Dim lr As Long
    lr = wsPetClick.Cells(wsPetClick.Rows.Count, "C").End(xlUp).Row
    If lr > 1 Then wsPetClick.Rows("2:" & lr).ClearContents

    lr = wsIruiClick.Cells(wsIruiClick.Rows.Count, "C").End(xlUp).Row
    If lr > 1 Then wsIruiClick.Rows("2:" & lr).ClearContents

    lr = wsSagawa.Cells(wsSagawa.Rows.Count, "C").End(xlUp).Row
    If lr > 1 Then wsSagawa.Rows("2:" & lr).ClearContents

    ' 非表示列を一時的にすべて表示（コピー漏れ防止）
    Dim lastEditCol As Integer: lastEditCol = 35  ' AI列まで
    Dim colHidden(1 To 35) As Boolean
    Dim c As Integer
    For c = 1 To lastEditCol
        colHidden(c) = wsEdit.Columns(c).Hidden
        wsEdit.Columns(c).Hidden = False
    Next c

    ' 最終行取得
    Dim lastRow As Long
    lastRow = wsEdit.Cells(wsEdit.Rows.Count, "C").End(xlUp).Row

    Dim nextRowPC As Long: nextRowPC = 2
    Dim nextRowIC As Long: nextRowIC = 2
    Dim nextRowSG As Long: nextRowSG = 2

    Dim i As Long
    For i = 2 To lastRow
        Dim shipMethod As String
        shipMethod = wsEdit.Cells(i, "A").Value

        Dim targetWs As Worksheet
        Dim nextRow As Long
        Dim category As String

        Select Case shipMethod
            Case "ペットクリック1通", "ペットクリック複数"
                Set targetWs = wsPetClick
                nextRow = nextRowPC
                category = "ペット"
                nextRowPC = nextRowPC + 1
            Case "衣類クリック1通", "衣類クリック複数"
                Set targetWs = wsIruiClick
                nextRow = nextRowIC
                category = "衣類"
                nextRowIC = nextRowIC + 1
            Case "ペット佐川", "衣類佐川"
                Set targetWs = wsSagawa
                nextRow = nextRowSG
                category = IIf(InStr(shipMethod, "ペット") > 0, "ペット", "衣類")
                nextRowSG = nextRowSG + 1
            Case Else
                GoTo NextRow
        End Select

        ' A列: 分類（ペット / 衣類）
        targetWs.Cells(nextRow, "A").Value = category

        ' C列〜AB列: Amazon注文データ26列をそのままコピー (order-id 〜 gift-message-text)
        targetWs.Cells(nextRow, "C").Resize(1, 26).Value = _
            wsEdit.Cells(i, "C").Resize(1, 26).Value

        ' AG列(編集) → AC列(貼付): is-prime
        targetWs.Cells(nextRow, "AC").Value = wsEdit.Cells(i, "AG").Value

        NextRow:
    Next i

    ' 非表示状態を元に戻す
    For c = 1 To lastEditCol
        wsEdit.Columns(c).Hidden = colHidden(c)
    Next c
End Sub


'--------------------------------------------------
' 実行3: 通知用CSVファイル作成（手順6〜8）
'   ※実行前に追跡番号レポートの貼り付けが必要
'--------------------------------------------------
Sub 実行3_Amazon_通知CSV作成()
    Application.ScreenUpdating = False

    Call 手順6_通知用データ作成
    Call 手順7_通知番号間違いチェック
    Call 手順8_通知用新規ファイル作成

    Application.ScreenUpdating = True

    MsgBox "通知用CSVを作成しました。" & Chr(13) & Chr(13) & _
           "Amazon管理画面の「発送完了報告データ」へ" & Chr(13) & _
           "CSVファイルをアップロードしてください。", _
           vbInformation, "実行3 完了"
End Sub
