Sub generateOutput()
    
    Dim inputSheet As Worksheet
    Dim outputSheet As Worksheet
    Dim lastRow As Long
    Dim manufacturer As String
    Dim vehicle As String
    Dim model As String
    Dim code As String
    Dim rowData As Variant
    Dim i As Long
    
    Set inputSheet = ActiveWorkbook.Sheets("入力シート")
    Set outputSheet = ActiveWorkbook.Sheets("出力シート")
    
    ' ユーザに入力シートをアクティブにするように求める
    MsgBox "入力シートをアクティブにしてください。", vbInformation, "情報"
    
    ' メーカ、車両、機種の3つの情報を入力させる
    manufacturer = InputBox("メーカを入力してください。", "メーカ入力")
    vehicle = InputBox("車両を入力してください。", "車両入力")
    model = InputBox("機種を入力してください。", "機種入力")
    
    ' データを処理する
    lastRow = inputSheet.Cells(Rows.Count, "B").End(xlUp).Row ' B列の最終行を取得
    For i = 1 To lastRow ' B列のデータを順に処理
        rowData = Split(inputSheet.Cells(i, "B"), "-") ' "-"で文字列を分割
        Select Case UBound(rowData) ' 分割された数によって場合分け
            Case 0 ' "-"が入っていない場合
                code = "(" & vehicle & "_" & manufacturer & "_" & model & ")-" & rowData(0)
            Case 1 ' "-"が1個入っている場合
                code = rowData(0) & "-(" & vehicle & "_" & manufacturer & "_" & model & ")-" & rowData(1) & "-00"
            Case 2 ' "-"が2個入っている場合
                code = rowData(0) & "-(" & vehicle & "_" & manufacturer & "_" & model & ")-" & rowData(1) & "-" & rowData(2) & "-00"
            Case Else ' "-"が3個以上入っている場合
                MsgBox "入力された文字列が不正です。処理を中止します。", vbCritical, "エラー"
                Exit Sub
        End Select
        outputSheet.Cells(i, "D") = code ' 処理結果をD列に出力
    Next i
    
    MsgBox "処理が完了しました。", vbInformation, "情報"
    
End Sub
