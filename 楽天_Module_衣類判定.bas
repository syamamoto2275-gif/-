Attribute VB_Name = "楽天_衣類判定"
'--------------------------------------------------
' 楽天fuu出荷自動化 - 衣類キーワード判定
' 商品名に衣類系キーワードが含まれていれば「衣類」と判定する。
' 楽天_マスター の 手順1 から呼び出される。
' 抽出元: 楽天fuu出荷データ作成用フォーマット2025_3_10.xlsm (2026-06-23更新版)
'--------------------------------------------------

Function IsClothing(productName As String) As Boolean
'
' 商品名に衣類系キーワードが含まれていれば True を返す。
' 衣類カテゴリの商品が増えたり、判定が誤っている場合は
' 下の keywords の一覧を編集してください(カンマで追加・削除するだけでOK)。
'
    Dim keywords As Variant
    Dim k As Variant

    keywords = Array("脇高", "セクシーレースリボン", "ブラジャー", "ブラ", "インナー")

    IsClothing = False
    For Each k In keywords
        If InStr(productName, CStr(k)) > 0 Then
            IsClothing = True
            Exit Function
        End If
    Next k
End Function
