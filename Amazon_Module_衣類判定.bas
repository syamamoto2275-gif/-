Attribute VB_Name = "Amazon_衣類判定"
'--------------------------------------------------
' Amazon出荷自動化 - 衣類キーワード判定
' 商品名に以下のキーワードが含まれれば「衣類」と判定する
'--------------------------------------------------

Function Is衣類Amazon(productName As String) As Boolean
    Dim keywords As Variant
    keywords = Array("脇高", "セクシー", "レース", "リボン", "ブラジャー", "ブラ", _
                     "インナー", "ショーツ", "下着", "ランジェリー", "パンティ", "ビキニ")
    Dim i As Integer
    For i = 0 To UBound(keywords)
        If InStr(productName, keywords(i)) > 0 Then
            Is衣類Amazon = True
            Exit Function
        End If
    Next i
    Is衣類Amazon = False
End Function
