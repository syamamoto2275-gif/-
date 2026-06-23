Attribute VB_Name = "Wowma_マスター"
'==================================================
' AUPay Market（Wowma）出荷自動化 - マスター実行モジュール
' ボタン3つで全工程を制御する
'
' 【ボタン配置】
'   実行1_Wowma_編集データ作成   → Sub 実行1_Wowma_編集データ作成
'   実行2_Wowma_出荷データ作成   → Sub 実行2_Wowma_出荷データ作成
'   実行3_Wowma_通知CSV作成      → Sub 実行3_Wowma_通知CSV作成
'==================================================

Function Is衣類Wowma(itemName As String) As Boolean
    Dim keywords As Variant
    keywords = Array("脇高", "セクシー", "レース", "リボン", "ブラジャー", "ブラ", _
                     "インナー", "ショーツ", "下着", "ランジェリー", "パンティ", "ビキニ")
    Dim i As Integer
    For i = 0 To UBound(keywords)
        If InStr(itemName, keywords(i)) > 0 Then
            Is衣類Wowma = True
            Exit Function
        End If
    Next i
    Is衣類Wowma = False
End Function


'--------------------------------------------------
' 実行1: 元データ → 編集シート作成 + 発送方法自動判定
'--------------------------------------------------
Sub 実行1_Wowma_編集データ作成()
    Application.ScreenUpdating = False
    Call 手順1_元データから編集データ作成
    Call Wowma_発送方法自動判定
    Application.ScreenUpdating = True
    MsgBox "編集データを作成しました。" & Chr(13) & Chr(13) & _
           "「編集」シートのA列（発送方法）を確認・修正してから" & Chr(13) & _
           "「実行2」ボタンを押してください。", vbInformation, "実行1 完了"
End Sub


'--------------------------------------------------
' 発送方法の自動判定（編集シートA列に書き込み）
'  G列 = itemName（商品名）
'  K列 = senderName（配送先氏名）
'  N列 = senderAddress（住所）
'--------------------------------------------------
Sub Wowma_発送方法自動判定()
    Dim ws As Worksheet
    Set ws = Sheets("編集")

    ' D列（orderId）で最終行を取得
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim i As Long
    For i = 2 To lastRow
        Dim itemName As String
        itemName = ws.Cells(i, "G").Value  ' G列 = itemName

        Dim senderName As String
        senderName = ws.Cells(i, "K").Value  ' K列 = senderName

        Dim address As String
        address = ws.Cells(i, "N").Value  ' N列 = senderAddress

        ' 同一送付先の件数
        Dim cnt As Long
        cnt = Application.WorksheetFunction.CountIfs( _
            ws.Columns("K"), senderName, _
            ws.Columns("N"), address)

        ' 衣類/ペット判定
        Dim isClothing As Boolean
        isClothing = Is衣類Wowma(itemName)

        ' A列に発送方法をセット
        If isClothing Then
            ws.Cells(i, "A").Value = IIf(cnt > 1, "衣類クリック複数", "衣類クリック1通")
        Else
            ws.Cells(i, "A").Value = IIf(cnt > 1, "ペットクリック複数", "ペットクリック1通")
        End If
    Next i
End Sub


'--------------------------------------------------
' 編集シートの各行を発送方法に応じて振り分ける
'--------------------------------------------------
Sub Wowma_振分_編集から貼付()
    Dim wsEdit As Worksheet
    Set wsEdit = Sheets("編集")

    Dim wsPetClick As Worksheet
    Dim wsIruiClick As Worksheet
    Dim wsSagawa As Worksheet
    Set wsPetClick = Sheets("ペット_クリック発行用_編集から貼付")
    Set wsIruiClick = Sheets("衣類_クリック発行用_編集から貼付")
    Set wsSagawa = Sheets("佐川発行用_編集から貼付")

    ' 各シートのデータをクリア（ヘッダー残す）
    Dim lr As Long
    lr = wsPetClick.Cells(wsPetClick.Rows.Count, "D").End(xlUp).Row
    If lr > 1 Then wsPetClick.Rows("2:" & lr).ClearContents

    lr = wsIruiClick.Cells(wsIruiClick.Rows.Count, "D").End(xlUp).Row
    If lr > 1 Then wsIruiClick.Rows("2:" & lr).ClearContents

    lr = wsSagawa.Cells(wsSagawa.Rows.Count, "D").End(xlUp).Row
    If lr > 1 Then wsSagawa.Rows("2:" & lr).ClearContents

    ' 非表示列を一時的に表示（コピー漏れ防止）
    Dim lastEditCol As Integer: lastEditCol = 30
    Dim colHidden(1 To 30) As Boolean
    Dim c As Integer
    For c = 1 To lastEditCol
        colHidden(c) = wsEdit.Columns(c).Hidden
        wsEdit.Columns(c).Hidden = False
    Next c

    Dim lastRow As Long
    lastRow = wsEdit.Cells(wsEdit.Rows.Count, "D").End(xlUp).Row

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

        ' A列: 発送方法、B列: 分類
        targetWs.Cells(nextRow, "A").Value = shipMethod
        targetWs.Cells(nextRow, "B").Value = category

        ' D列〜（Wowmaデータ: controlType〜orderStatus = 22列）
        targetWs.Cells(nextRow, "D").Resize(1, 22).Value = _
            wsEdit.Cells(i, "C").Resize(1, 22).Value

        NextRow:
    Next i

    ' 非表示状態を元に戻す
    For c = 1 To lastEditCol
        wsEdit.Columns(c).Hidden = colHidden(c)
    Next c
End Sub


'--------------------------------------------------
' 実行2: 出荷データ作成一式（手順2〜5）
'--------------------------------------------------
Sub 実行2_Wowma_出荷データ作成()
    Application.ScreenUpdating = False
    Call Wowma_振分_編集から貼付
    Call 手順2_物出しリスト作成_並べ替え
    Call 手順3_ペット_クリック発行用_一括発行アップデータ作成
    Call 手順4_衣類_クリック発行用_一括発行アップデータ作成
    Call 手順5_佐川発行用_E飛伝アップデータ作成
    Application.ScreenUpdating = True
    MsgBox "出荷データを作成しました。" & Chr(13) & Chr(13) & _
           "1. クリックポストで宛名を発行する" & Chr(13) & _
           "2. e飛伝で佐川伝票を印刷する" & Chr(13) & _
           "3. 追跡番号レポートを各シートに貼り付ける" & Chr(13) & Chr(13) & _
           "完了したら「実行3」を押してください。", vbInformation, "実行2 完了"
End Sub


'--------------------------------------------------
' 実行3: 通知用データ作成（手順6〜8）
'--------------------------------------------------
Sub 実行3_Wowma_通知CSV作成()
    Application.ScreenUpdating = False
    Call 手順6_追跡番号を通知用データに貼付
    Call 手順7_通知番号間違いチェック
    Call 手順8_通知用新規ファイル作成
    Application.ScreenUpdating = True
    MsgBox "通知用ファイルを作成しました。" & Chr(13) & _
           "AUPay Marketの管理画面にアップロードしてください。", _
           vbInformation, "実行3 完了"
End Sub
