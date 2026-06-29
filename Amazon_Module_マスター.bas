Attribute VB_Name = "Amazon_マスター"
'==================================================
' Amazon出荷自動化 - マスター実行モジュール（実物2026.6.17版を基準）
' ボタン3つで全工程を制御する
'
' 【ボタン配置】
'   実行1_Amazon_編集データ作成   → Sub 実行1_Amazon_編集データ作成
'   実行2_Amazon_出荷データ作成   → Sub 実行2_Amazon_出荷データ作成
'   実行3_Amazon_通知CSV作成      → Sub 実行3_Amazon_通知CSV作成
'
' 2026-06-29 修正（社長指示）:
'   1) 実行1後、編集シートの表示倍率を100%に戻す
'   2) 実行2で佐川の該当注文が無いときは空ファイルを作らない
'   3) 実行2後、ペットクリック物出しシートを表示する
'   4) ペットクリック物出しシートの印刷をA4横に設定
'   5) 各「編集から貼付」シートのA列(分類)に、正確な発送方法を表示
'   6) 住所(T:V)の文字化け・日付変換を防ぐため表示文字列で転記
'==================================================

Sub 実行1_Amazon_編集データ作成()
    Application.ScreenUpdating = False
    Call 手順1_元データから編集データ作成
    Call Amazon_発送方法自動判定
    Application.ScreenUpdating = True

    ' (1) 編集シートを表示し、ズームを100%に戻す
    Sheets("編集").Select
    ActiveWindow.Zoom = 100

    MsgBox "編集データを作成しました。" & Chr(13) & Chr(13) & _
           "次のステップ:" & Chr(13) & _
           "「編集」シートのA列（発送方法）を確認・修正してください。" & Chr(13) & _
           "コスト面で佐川の方が良い場合は、「ペット佐川」「衣類佐川」に変更できます。" & Chr(13) & Chr(13) & _
           "確認が完了したら「実行2」ボタンを押してください。", _
           vbInformation, "実行1 完了"
End Sub

Sub Amazon_発送方法自動判定()
    Dim ws As Worksheet
    Set ws = Sheets("編集")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    Dim i As Long
    For i = 2 To lastRow
        Dim productName As String
        productName = ws.Cells(i, "N").Value
        Dim recipientName As String
        recipientName = ws.Cells(i, "S").Value
        Dim address1 As String
        address1 = ws.Cells(i, "T").Value
        Dim cnt As Long
        cnt = Application.WorksheetFunction.CountIfs( _
            ws.Columns("S"), recipientName, _
            ws.Columns("T"), address1)
        Dim isClothing As Boolean
        isClothing = Is衣類Amazon(productName)
        If isClothing Then
            If cnt >= 7 Then
                ws.Cells(i, "A").Value = "衣類佐川"
            ElseIf cnt > 3 Then
                ws.Cells(i, "A").Value = "衣類クリック複数"
            Else
                ws.Cells(i, "A").Value = "衣類クリック1通"
            End If
        Else
            Dim isPetClick As Boolean
            isPetClick = False
            If InStr(productName, "ごはん") > 0 Or InStr(productName, "おやつ") > 0 Then
                If InStr(productName, "3kg") = 0 And InStr(productName, "5kg") = 0 And InStr(productName, "10kg") = 0 And _
                   InStr(productName, "3ｋｇ") = 0 And InStr(productName, "5ｋｇ") = 0 And InStr(productName, "10ｋｇ") = 0 Then
                    isPetClick = True
                End If
            End If
            If isPetClick Then
                ws.Cells(i, "A").Value = IIf(cnt > 1, "ペットクリック複数", "ペットクリック1通")
            Else
                ws.Cells(i, "A").Value = "ペット佐川"
            End If
        End If
    Next i
End Sub

Sub 実行2_Amazon_出荷データ作成()
    Application.ScreenUpdating = False
    Call Amazon_振分_編集から貼付
    Call 手順2_物出しリスト作成
    Call 手順3_ペット_クリック発行用_一括発行アップデータ
    Call 手順4_衣類_クリック発行用_一括発行アップデータ

    ' (2) 佐川の該当注文が無ければ、空のE飛伝ファイルを作らない
    If Sheets("佐川発行用_編集から貼付").Cells(2, "C").Value <> "" Then
        Call 手順5_佐川発行用_E飛伝アップデータ
    End If

    ' (4) ペットクリック物出しシートの印刷をA4横に設定
    With Sheets("ペットクリック物出し").PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With

    Application.ScreenUpdating = True

    ' (3) ペットクリック物出しシートを表示する
    Sheets("ペットクリック物出し").Select
    Range("A1").Select

    MsgBox "出荷データを作成しました。" & Chr(13) & Chr(13) & _
           "次のステップ:" & Chr(13) & _
           "1. クリックポストで宛名を発行する" & Chr(13) & _
           "2. e飛伝で佐川伝票を印刷する" & Chr(13) & _
           "3. 追跡番号レポートを各シートに貼り付ける" & Chr(13) & Chr(13) & _
           "完了したら「実行3」ボタンを押してください。", _
           vbInformation, "実行2 完了"
End Sub

Sub Amazon_振分_編集から貼付()
    Dim wsEdit As Worksheet
    Set wsEdit = Sheets("編集")
    Dim wsPetClick As Worksheet
    Dim wsIruiClick As Worksheet
    Dim wsSagawa As Worksheet
    Set wsPetClick = Sheets("ペット_クリック発行用_編集から貼付")
    Set wsIruiClick = Sheets("衣類_クリック発行用_編集から貼付")
    Set wsSagawa = Sheets("佐川発行用_編集から貼付")
    Dim lr As Long
    lr = wsPetClick.Cells(wsPetClick.Rows.Count, "C").End(xlUp).Row
    If lr > 1 Then wsPetClick.Rows("2:" & lr).ClearContents
    lr = wsIruiClick.Cells(wsIruiClick.Rows.Count, "C").End(xlUp).Row
    If lr > 1 Then wsIruiClick.Rows("2:" & lr).ClearContents
    lr = wsSagawa.Cells(wsSagawa.Rows.Count, "C").End(xlUp).Row
    If lr > 1 Then wsSagawa.Rows("2:" & lr).ClearContents
    ' 住所列(T:V)を文字列形式に設定（番地の日付変換防止）
    wsPetClick.Columns("T:V").NumberFormat = "@"
    wsIruiClick.Columns("T:V").NumberFormat = "@"
    wsSagawa.Columns("T:V").NumberFormat = "@"
    Dim lastEditCol As Integer: lastEditCol = 35
    Dim colHidden(1 To 35) As Boolean
    Dim c As Integer
    For c = 1 To lastEditCol
        colHidden(c) = wsEdit.Columns(c).Hidden
        wsEdit.Columns(c).Hidden = False
    Next c
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
        Dim NextRow As Long
        Select Case shipMethod
            Case "ペットクリック1通", "ペットクリック複数"
                Set targetWs = wsPetClick
                NextRow = nextRowPC
                nextRowPC = nextRowPC + 1
            Case "衣類クリック1通", "衣類クリック複数"
                Set targetWs = wsIruiClick
                NextRow = nextRowIC
                nextRowIC = nextRowIC + 1
            Case "ペット佐川", "衣類佐川"
                Set targetWs = wsSagawa
                NextRow = nextRowSG
                nextRowSG = nextRowSG + 1
            Case Else
                GoTo NextRow
        End Select
        ' (5) A列(分類)に正確な発送方法をそのまま表示（例:ペットクリック1通/複数）
        targetWs.Cells(NextRow, "A").Value = shipMethod
        targetWs.Cells(NextRow, "C").Resize(1, 26).Value = _
            wsEdit.Cells(i, "C").Resize(1, 26).Value
        ' (6) 住所(T:V)は文字化け・日付変換を防ぐため、表示文字列をそのまま転記
        targetWs.Cells(NextRow, "T").Value = wsEdit.Cells(i, "T").Text
        targetWs.Cells(NextRow, "U").Value = wsEdit.Cells(i, "U").Text
        targetWs.Cells(NextRow, "V").Value = wsEdit.Cells(i, "V").Text
        targetWs.Cells(NextRow, "AC").Value = wsEdit.Cells(i, "AG").Value
NextRow:
    Next i
    For c = 1 To lastEditCol
        wsEdit.Columns(c).Hidden = colHidden(c)
    Next c
End Sub

Sub 実行3_Amazon_通知CSV作成()
    Application.ScreenUpdating = False
    Call 手順6_通知用データ作成
    Call 手順7_通知番号間違いチェック
    Call 手順8_通知用新規ファイル作成
    Application.ScreenUpdating = True
    MsgBox "通知用CSVを作成しました。" & Chr(13) & Chr(13) & _
           "Amazon管理画面の「発送完了通知データ」へ" & Chr(13) & _
           "CSVファイルをアップロードしてください。", _
           vbInformation, "実行3 完了"
End Sub
