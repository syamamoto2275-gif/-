Attribute VB_Name = "楽天_マスター"
'==================================================
' 楽天fuu出荷自動化 - マスター実行モジュール
' ボタン3つで全工程を制御する
'
' 【ボタン配置】
'   実行1_編集データ作成        → Sub 実行1_編集データ作成
'   実行2_出荷データ作成一式     → Sub 実行2_出荷データ作成一式
'   実行3_通知CSV作成           → Sub 実行3_通知CSV作成
'
' 衣類判定は別モジュール「楽天_衣類判定」の Function IsClothing を使用する。
' 抽出元: 楽天fuu出荷データ作成用フォーマット2025_3_10.xlsm (2026-06-23更新版)
'==================================================


'--------------------------------------------------
' 実行1: 元データ → 編集シート作成 + 発送方法自動判定
'   ここで一度「編集」シートA列(発送方法)を確認・修正してから次へ進む
'--------------------------------------------------
Sub 実行1_編集データ作成()
'
' 手順1(元データから編集データ作成、発送方法の自動判定)のみを実行する。
' ボタンを分けているのは、ここで一度「編集」シートA列(発送方法)を
' 確認・修正してから次に進めるようにするため。
' 例:3kgの商品を福岡へ送る場合、佐川急便なら770円だが
'     クリックポスト185円×3通(複数)の方が安く済む、といったケースを
'     ここで手動で「ペット佐川」→「ペットクリック複数」等に変更できる。
'
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    手順1_元データから編集データ作成

    Application.ScreenUpdating = True
    MsgBox "編集データを作成しました。" & vbCrLf & vbCrLf & _
           "「編集」シートのA列(発送方法)を確認してください。" & vbCrLf & _
           "送料を抑えるために変更したい行があれば、ドロップダウンから選び直してから" & vbCrLf & _
           "「実行2_出荷データ作成一式」を実行してください。", vbInformation, "編集データ作成 完了"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "処理中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical, "エラー"
End Sub

'--------------------------------------------------
' 実行2: 物出しリスト作成 → 発行用貼付シート自動作成 →
'        ペットクリック/衣類クリック/佐川 の発行用データ作成
'--------------------------------------------------
Sub 実行2_出荷データ作成一式()
'
' 手順2~5(物出しリスト作成→発行用貼付シート自動作成→
' ペットクリック/衣類クリック/佐川の発行用データ作成)をまとめて1クリックで実行する。
' 「編集」シートA列の発送方法の確認・修正が済んでから実行すること。
'
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    手順2_物出しリスト作成
    手順2b_発行用貼付シート自動作成
    手順3_ペット_クリック発行用_一括発行アップデータ
    手順4_衣類_クリック発行用_一括発行アップデータ
    手順5_佐川発行用_E飛伝アップデータ

    Application.ScreenUpdating = True

    ' (楽天-1) ペットクリック物出しシートを表示する
    Sheets("ペットクリック物出し").Select
    Range("A1").Select

    MsgBox "出荷データの作成が完了しました。" & vbCrLf & vbCrLf & _
           "クリックポスト・佐川それぞれの発行用データができています。" & vbCrLf & _
           "各システムでラベルを発行し、追跡番号レポートを貼り付けたら、" & vbCrLf & _
           "「実行3_通知CSV作成」ボタンを押してください。", vbInformation, "出荷データ作成 完了"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "処理中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical, "エラー"
End Sub

'--------------------------------------------------
' 実行3: 通知用データ作成 → 通知用新規CSVファイル作成
'   事前に各追跡番号レポートを貼付シートに貼り付けておくこと
'--------------------------------------------------
Sub 実行3_通知CSV作成()
'
' 手順6~7をまとめて1クリックで実行する。
' 事前に「ペット_クリック追跡番号レポート貼付」「衣類_クリック追跡番号レポート貼付」
' シートに、各システムから出力した追跡番号レポートを貼り付けておくこと。
'
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    手順6_通知用データ作成
    手順7_通知用新規CSVファイル作成

    Application.ScreenUpdating = True
    MsgBox "通知用CSVファイルの作成が完了しました。", vbInformation, "通知CSV作成 完了"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "処理中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical, "エラー"
End Sub


'==================================================
' 以下、各手順マクロ（実行ボタンから呼び出される個別処理）
'==================================================

Sub 手順1_元データから編集データ作成()
'
' 手順1_元データから編集データ作成 Macro (発送方法 自動判定版)
' 変更点: これまで「編集」シートA列にドロップダウンで手入力していた
'         発送方法(ペットクリック1通/複数、衣類クリック1通/複数、ペット佐川、衣類佐川)を、
'         配送方法(メール便/宅配便)・商品名のキーワード・同送付先の件数から自動判定する。
'         ドロップダウン自体は手動で修正したい場合のために残してある。
'
    Dim lastRow As Long, i As Long
    Dim shippingMethod As String, productName As String
    Dim category As String, methodPart As String, countPart As String
    Dim cnt As Long

    Sheets("編集").Select
    Cells.Clear

    Sheets("元データ").Select
    Cells.Select
    Selection.Copy
    Sheets("編集").Select
    Range("A1").Select
    ActiveSheet.Paste
    Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "発送方法"
    Columns("A:A").ColumnWidth = 15.5
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "複数"

    lastRow = Cells(Rows.Count, "D").End(xlUp).Row  ' D列=注文番号で実データ行数を判定
    If lastRow < 2 Then
        MsgBox "元データに処理対象のデータがありません。", vbExclamation
        Exit Sub
    End If

    ' B列: 同一送付先(送付先姓+送付先名+都道府県+郡市区)の行数をカウント(既存ロジックのまま)
    Range("B2:B" & lastRow).FormulaR1C1 = "=COUNTIFS(C10,RC[8],C11,RC[9],C16,RC[14],C17,RC[15])"
    ' A列: 発送方法を自動判定
    ' W列=商品名、B列=同一送付先件数
    For i = 2 To lastRow
        productName = CStr(Cells(i, "W").Value)
        cnt = Cells(i, "B").Value

        If IsClothing(productName) Then
            If cnt >= 7 Then
                Cells(i, "A").Value = "衣類佐川"
            ElseIf cnt > 3 Then
                Cells(i, "A").Value = "衣類クリック複数"
            Else
                Cells(i, "A").Value = "衣類クリック1通"
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
                Cells(i, "A").Value = IIf(cnt > 1, "ペットクリック複数", "ペットクリック1通")
            Else
                Cells(i, "A").Value = "ペット佐川"
            End If
        End If
    Next i

    ' 自動判定が間違っている場合に手動で直せるよう、ドロップダウンも残しておく
    Range("A2:A" & lastRow).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="ペットクリック1通,ペットクリック複数,衣類クリック1通,衣類クリック複数,ペット佐川,衣類佐川"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With

    Columns("E:I").Select
    Selection.EntireColumn.Hidden = True
    Columns("L:O").Select
    Selection.EntireColumn.Hidden = True
    Columns("Q:V").Select
    Selection.EntireColumn.Hidden = True
    Columns("X:Y").Select
    Selection.EntireColumn.Hidden = True
    Columns("W:W").EntireColumn.AutoFit
    Columns("Z:Z").EntireColumn.AutoFit
    Columns("AA:AA").EntireColumn.AutoFit

    Range("A2:AF600").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$B2=2"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
    End With
    Range("A2:AF600").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$B2=3"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 15773696
        .TintAndShade = 0
    End With
    Range("A2:AF600").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$B2=4"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
    End With
    Range("A2:AF600").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$B2=5"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
    End With
    Range("A2:AF600").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$B2=6"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    Range("J2:Z600").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$Z2>1"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .TintAndShade = 0
    End With
End Sub


Sub 手順2_物出しリスト作成()
'
' 手順2_物出しリスト作成 Macro
'
    Sheets("編集").Select
    Columns("A:A").Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("編集").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("編集").Sort.SortFields.Add Key:=Range("A1"), SortOn _
        :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("編集").Sort
        .SetRange Range("A2:AF600")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("編集").Select
    Cells.Select
    Selection.Copy
    Sheets("衣類物出し").Select
    Range("A1").Select
    ActiveSheet.Paste
    Range("J1:AA600").Select
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    Range("$A$1:$A$600").AutoFilter Field:=1, Criteria1:=Array( _
        "衣類クリック1通", "衣類クリック複数", "衣類佐川"), Operator:=xlFilterValues
    Columns("P:P").Select
    Selection.EntireColumn.Hidden = True
    Columns("J:AA").Select
    ActiveSheet.PageSetup.PrintArea = "$J:$AA"
    Range("A1").Select
    Sheets("編集").Select
    Cells.Select
    Selection.Copy
    Sheets("ペット佐川物出し").Select
    Range("A1").Select
    ActiveSheet.Paste
    Range("J1:AA600").Select
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    Range("$A$1:$A$600").AutoFilter Field:=1, Criteria1:=Array( _
        "ペット佐川"), Operator:=xlFilterValues
    Columns("P:P").Select
    Selection.EntireColumn.Hidden = True
    Columns("J:AA").Select
    ActiveSheet.PageSetup.PrintArea = "$J:$AA"
    Range("A1").Select
    Sheets("編集").Select
    Cells.Select
    Selection.Copy
    Sheets("ペットクリック物出し").Select
    Range("A1").Select
    ActiveSheet.Paste
    Range("J1:AA600").Select
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    Range("$A$1:$A$600").AutoFilter Field:=1, Criteria1:=Array( _
        "ペットクリック1通", "ペットクリック複数"), Operator:=xlFilterValues
    ' (楽天-2) フィルター済みのペットクリック物出しを商品名(W列)で並び替える
    ActiveWorkbook.Worksheets("ペットクリック物出し").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ペットクリック物出し").Sort.SortFields.Add _
        Key:=Range("W1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ペットクリック物出し").Sort
        .SetRange Range("A1:AA600")
        .Header = xlYes
        .Orientation = xlTopToBottom
        .Apply
    End With
    Columns("P:P").Select
    Selection.EntireColumn.Hidden = True
    Columns("J:AA").Select
    ActiveSheet.PageSetup.PrintArea = "$J:$AA"
    Range("A1").Select

End Sub


Sub 手順2b_発行用貼付シート自動作成()
'
' 「編集」シートのデータを発送方法に応じて、
' ペット_クリック発行用_編集から貼付 / 衣類_クリック発行用_編集から貼付 / 佐川発行用_編集から貼付
' へ自動的に振り分けてコピーする。
' (これまで手作業でコピー&ペーストしていた部分を自動化)
'
    DistributeByCategory "ペット_クリック発行用_編集から貼付", Array("ペットクリック1通", "ペットクリック複数")
    DistributeByCategory "衣類_クリック発行用_編集から貼付", Array("衣類クリック1通", "衣類クリック複数")
    DistributeByCategory "佐川発行用_編集から貼付", Array("ペット佐川", "衣類佐川")
End Sub

Sub DistributeByCategory(targetSheet As String, methods As Variant)
'
' 「編集」シートをmethodsで指定した発送方法だけにフィルタし、
' 表示されている行だけをtargetSheetのA1に貼り付ける(行を詰めてコピー)。
'
    Dim lastRow As Long, lastCol As Long

    Sheets("編集").Select
    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilterMode = False

    lastRow = Cells(Rows.Count, "D").End(xlUp).Row
    lastCol = Cells(1, Columns.Count).End(xlToLeft).Column

    Range(Cells(1, 1), Cells(lastRow, lastCol)).AutoFilter Field:=1, Criteria1:=methods, Operator:=xlFilterValues

    Sheets(targetSheet).Cells.Clear

    Range(Cells(1, 1), Cells(lastRow, lastCol)).Select
    Selection.Copy
    Sheets(targetSheet).Select
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

    Sheets("編集").AutoFilterMode = False
End Sub


Sub 手順3_ペット_クリック発行用_一括発行アップデータ()
'
' 手順3_ペット_クリック発行用_一括発行アップデータ Macro
'
    Sheets("ペット_クリック発行用_一括発行アップデータ").Select
    Range("A2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=RIGHT(""000"" & ペット_クリック発行用_編集から貼付!RC[13],3)&""-""&RIGHT(""0000"" & ペット_クリック発行用_編集から貼付!RC[14],4)"
    Range("A2").Select
    Selection.AutoFill Destination:=Range("A2:A600"), Type:=xlFillDefault
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=ペット_クリック発行用_編集から貼付!RC[8]&ペット_クリック発行用_編集から貼付!RC[9]"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B600")
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=ペット_クリック発行用_編集から貼付!RC[12]"
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D600")
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=ペット_クリック発行用_編集から貼付!RC[12]"
    Range("E2").Select
    Selection.AutoFill Destination:=Range("E2:E600")
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=ペット_クリック発行用_編集から貼付!RC[12]"
    Range("F2").Select
    Selection.AutoFill Destination:=Range("F2:F600")

End Sub


Sub 手順4_衣類_クリック発行用_一括発行アップデータ()
'
' 手順4_衣類_クリック発行用_一括発行アップデータ Macro
'
    Sheets("衣類_クリック発行用_一括発行アップデータ").Select
    Range("A2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=RIGHT(""000"" & 衣類_クリック発行用_編集から貼付!RC[13],3)&""-""&RIGHT(""0000"" & 衣類_クリック発行用_編集から貼付!RC[14],4)"
    Range("A2").Select
    Selection.AutoFill Destination:=Range("A2:A600"), Type:=xlFillDefault
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=衣類_クリック発行用_編集から貼付!RC[8]&衣類_クリック発行用_編集から貼付!RC[9]"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B600")
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=衣類_クリック発行用_編集から貼付!RC[12]"
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D600")
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=衣類_クリック発行用_編集から貼付!RC[12]"
    Range("E2").Select
    Selection.AutoFill Destination:=Range("E2:E600")
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=衣類_クリック発行用_編集から貼付!RC[12]"
    Range("F2").Select
    Selection.AutoFill Destination:=Range("F2:F600")
End Sub


Sub 手順5_佐川発行用_E飛伝アップデータ()
'
' 手順5_佐川発行用_E飛伝アップデータ Macro
'
    Sheets("佐川発行用_E飛伝アップデータ").Select
    Range("C2").Select
    ActiveCell.FormulaR1C1 = _
        "=TEXT(佐川発行用_編集から貼付!RC[16],""000"")&TEXT(佐川発行用_編集から貼付!RC[17],""0000"")&TEXT(佐川発行用_編集から貼付!RC[18],""0000"")"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C200"), Type:=xlFillDefault
    Sheets("佐川発行用_E飛伝アップデータ").Select
    Range("D2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=RIGHT(""000"" & 佐川発行用_編集から貼付!RC[10],3)&""-""&RIGHT(""0000"" & 佐川発行用_編集から貼付!RC[11],4)"
        Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D200"), Type:=xlFillDefault
    Sheets("佐川発行用_編集から貼付").Select
    Range("P2:P200").Select
    Selection.Copy
    Sheets("佐川発行用_E飛伝アップデータ").Select
    Range("E2:E200").Select
    ActiveSheet.Paste
    Sheets("佐川発行用_編集から貼付").Select
    Range("Q2:Q200").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("佐川発行用_E飛伝アップデータ").Select
    Range("F2:F200").Select
    ActiveSheet.Paste
    Sheets("佐川発行用_編集から貼付").Select
    Range("R2:R200").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("佐川発行用_E飛伝アップデータ").Select
    Range("G2:G200").Select
    ActiveSheet.Paste
    Range("H2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=佐川発行用_編集から貼付!RC[2]&"" ""&佐川発行用_編集から貼付!RC[3]"
    Range("H2").Select
    Selection.AutoFill Destination:=Range("H2:H200"), Type:=xlFillDefault
    Sheets("佐川発行用_編集から貼付").Select
    Range("D2:D200").Select
    Selection.Copy
    Sheets("佐川発行用_E飛伝アップデータ").Select
    Range("Z2:Z200").Select
    ActiveSheet.Paste

    ' (楽天-3) お届け先住所(E:G列)は太字にしない
    Sheets("佐川発行用_E飛伝アップデータ").Range("E2:G200").Font.Bold = False

    Sheets("佐川発行用_E飛伝アップデータ").Select
    Sheets("佐川発行用_E飛伝アップデータ").Copy
    n = Cells(Rows.Count, "E").End(xlUp).Row + 1
    Range("E" & n).Select
    Rows(ActiveCell.Row & ":" & Rows.Count).Delete

    y = Year(Date)
    m = Format(Month(Date), "00")
    d = Format(Day(Date), "00")
    Sheets("佐川発行用_E飛伝アップデータ").Select

    Dim userName As String
    userName = Environ("USERNAME")
    Dim filePath As String
    filePath = "C:\Users\" & userName & "\Dropbox\ネットショップ\【出荷】fuu楽天\fuu楽天_佐川E飛伝\"

        ActiveWorkbook.SaveAs Filename:=filePath & "fuu楽天_佐川E飛伝" & "_" & y & "_" & m & d & ".csv" _
        , FileFormat:=xlCSV, CreateBackup:=False

    ActiveWorkbook.Close
End Sub


Sub 手順6_通知用データ作成()
'
' 手順6_通知用データ作成 Macro
'
    Sheets("ペット_クリック追跡番号レポート貼付").Select
    Range("C2:C600").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("通知用データ").Select
    Range("E2").Select
    ActiveSheet.Paste
    Sheets("ペット_クリック追跡番号レポート貼付").Select
    Range("E2:E600").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("通知用データ").Select
    Range("D2").Select
    ActiveSheet.Paste

    Sheets("衣類_クリック追跡番号レポート貼付").Select
    Range("C2:C600").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("通知用データ").Select
    n = Cells(Rows.Count, "E").End(xlUp).Row + 1
    Range("E" & n).Select
    ActiveSheet.Paste
    Sheets("衣類_クリック追跡番号レポート貼付").Select
    Range("E2:E600").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("通知用データ").Select
    n = Cells(Rows.Count, "D").End(xlUp).Row + 1
    Range("D" & n).Select
    ActiveSheet.Paste

    Sheets("編集").Select
    Range("D2:E600").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("通知用データ").Select
    Range("A2").Select
    ActiveSheet.Paste
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=編集!RC[7]&編集!RC[8]"
    Selection.AutoFill Destination:=Range("C2:C600"), Type:=xlFillDefault
    Range("A2:C600").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A2").Select

End Sub


Sub 手順7_通知用新規CSVファイル作成()
'
' 手順7_通知用新規CSVファイル作成 Macro
'
    Sheets("通知用データ").Select
    Sheets("通知用データ").Copy
    Range("C2:C600").Select
    Selection.ClearContents
    Columns("I:J").Select
    Selection.Delete Shift:=xlToLeft
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
    n = Cells(Rows.Count, "A").End(xlUp).Row + 1
    Range("A" & n).Select
    Rows(ActiveCell.Row & ":" & Rows.Count).Delete

    y = Year(Date)
    m = Format(Month(Date), "00")
    d = Format(Day(Date), "00")
    Sheets("通知用データ").Select

    Dim userName As String
    userName = Environ("USERNAME")
    Dim filePath As String
    filePath = "C:\Users\" & userName & "\Dropbox\ネットショップ\【出荷】fuu楽天\fuu楽天_発送通知\"

        ActiveWorkbook.SaveAs Filename:=filePath & "fuu_Rakuten" & "_" & y & "_" & m & d & ".csv" _
        , FileFormat:=xlCSV, CreateBackup:=False

    ActiveWorkbook.Close
End Sub
