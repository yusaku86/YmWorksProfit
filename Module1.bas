Attribute VB_Name = "Module1"
'// YM売掛金雛形作成モジュール
Option Explicit

'// メインプロシージャ
Public Sub main()

    Call calculateTotalAmount

    Sheets("YM売上雛形").Activate

    Dim rs As New ADODB.Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM [ワーク$]", connectDb(ThisWorkbook.FullName), adOpenStatic, adLockOptimistic

    rs.Sort = "請求先任意CD ASC"

    Dim targetRow As Long: targetRow = 4

    '// ﾘｻｲｸﾙ料の合計
    Dim totalRecyclingCharge
    
    Do Until rs.EOF
        If rs!請求先任意CD = 0 Or rs!請求先任意CD = 5013 Or rs!請求先任意CD = 1121 Or rs!請求先任意CD = 1273 Or rs!請求先任意CD = 1166 Then
            GoTo DoContinue
        End If

        '// 課税金額入力
        Cells(targetRow, 1).Value = rs!売上区分CD
        Cells(targetRow, 2).Value = rs!請求先任意CD
        Cells(targetRow, 3).Value = rs!請求先名1
        Cells(targetRow, 4).Value = rs!課税小計
        Cells(targetRow, 5).Value = rs!消費税計
        Cells(targetRow, 14).Value = rs!課税小計 + rs!消費税計
        
        Cells(targetRow, 15).Value = rs!車両登録番号支局名 & rs!車両登録番号分類 & rs!車両登録番号記号 & rs!車両登録番号番号 & " " & rs!売上区分名称
        
        '// 自賠責金額が0円より大きい場合
        If rs!自賠責金額 > 0 Then
            targetRow = targetRow + 1
            Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
            
            Cells(targetRow, 6).Value = rs!自賠責金額
            Cells(targetRow, 14).Value = rs!自賠責金額
        
            Cells(targetRow, 15).Value = rs!車両登録番号支局名 & rs!車両登録番号分類 & rs!車両登録番号記号 & rs!車両登録番号番号 & " " & rs!売上区分名称
        End If
        
        '// 重量税金額が0円より大きい場合
        If rs!重量税金額 > 0 Then
            targetRow = targetRow + 1
            Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
        
            Cells(targetRow, 7).Value = rs!重量税金額
            Cells(targetRow, 14).Value = rs!重量税金額
                    
            Cells(targetRow, 15).Value = rs!車両登録番号支局名 & rs!車両登録番号分類 & rs!車両登録番号記号 & rs!車両登録番号番号 & " " & rs!売上区分名称
        End If
        
        '/**
         '* 諸費用金額が0円より大きい場合
         '* ※リサイクル料はシートに表示しないが、合計金額が必要のため合計する
        '**/
        Dim i As Long
        
        For i = 1 To 5
            If rs.Fields("諸費用金額" & i).Value <= 0 Then
                GoTo ForContinue
            End If
            
            '// 諸費用金額1から諸費用金額5まで金額を確認し、種類によって入力するマスを変更する
            Select Case rs.Fields("諸費用名称" & i).Value
            
                Case "リサイクル料"
                    totalRecyclingCharge = totalRecyclingCharge + rs.Fields("諸費用金額" & i)
                    GoTo ForContinue
            
                Case "検査登録印紙"
                    targetRow = targetRow + 1
                    Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
                    
                    Cells(targetRow, 8).Value = rs.Fields("諸費用金額" & i).Value
                    Cells(targetRow, 15).Value = rs!車両登録番号支局名 & rs!車両登録番号分類 & rs!車両登録番号記号 & rs!車両登録番号番号 & " " & rs!売上区分名称
                
                Case "車検印紙代"
                    targetRow = targetRow + 1
                    Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
                    
                    Cells(targetRow, 9).Value = rs.Fields("諸費用金額" & i).Value
                    Cells(targetRow, 15).Value = rs!車両登録番号支局名 & rs!車両登録番号分類 & rs!車両登録番号記号 & rs!車両登録番号番号 & " " & rs!売上区分名称
                
                Case "臨時運行許可証"
                    targetRow = targetRow + 1
                    Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
                    
                    Cells(targetRow, 10).Value = rs.Fields("諸費用金額" & i).Value
                    Cells(targetRow, 15).Value = rs!車両登録番号支局名 & rs!車両登録番号分類 & rs!車両登録番号記号 & rs!車両登録番号番号 & " " & rs!売上区分名称
                
                Case "登録番号標"
                    targetRow = targetRow + 1
                    Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
                    
                    Cells(targetRow, 11).Value = rs.Fields("諸費用金額" & i).Value
                    Cells(targetRow, 15).Value = rs!車両登録番号支局名 & rs!車両登録番号分類 & rs!車両登録番号記号 & rs!車両登録番号番号 & " " & rs!売上区分名称
                
                Case "車両番号標"
                    targetRow = targetRow + 1
                    Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
                    
                    Cells(targetRow, 12).Value = rs.Fields("諸費用金額" & i).Value
                    Cells(targetRow, 15).Value = rs!車両登録番号支局名 & rs!車両登録番号分類 & rs!車両登録番号記号 & rs!車両登録番号番号 & " " & rs!売上区分名称
                
                Case "自動車税種別割"
                    targetRow = targetRow + 1
                    Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
                    
                    Cells(targetRow, 13).Value = rs.Fields("諸費用金額" & i).Value
                    Cells(targetRow, 15).Value = rs!車両登録番号支局名 & rs!車両登録番号分類 & rs!車両登録番号記号 & rs!車両登録番号番号 & " " & rs!売上区分名称
            End Select
            
            Cells(targetRow, 14).Value = rs.Fields("諸費用金額" & i).Value
    
ForContinue:
        Next
    
        targetRow = targetRow + 1
DoContinue:
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    '// リサイクル料の合計入力
    Cells(2, 7).Value = totalRecyclingCharge
    
    '// 合計のセルに式入力
    Cells(2, 5).Formula = "=SUM(N:N)"
    
End Sub

'// 総計、コード0(山岸運送)の合計、社内間合計を計算
Private Sub calculateTotalAmount()

    '// 総計
    Cells(1, 5).Value = WorksheetFunction.Sum(Range(Sheets("ワーク").Cells(2, 111), Sheets("ワーク").Cells(Rows.Count, 111).End(xlUp)))
    
    '// コード0
    Cells(1, 7).Value = WorksheetFunction.SumIf(Sheets("ワーク").Columns(24), 0, Sheets("ワーク").Columns(111))
    
    '// 社内間
    Cells(1, 12).Value = _
        WorksheetFunction.SumIf(Sheets("ワーク").Columns(24), 5013, Sheets("ワーク").Columns(111)) _
        + WorksheetFunction.SumIf(Sheets("ワーク").Columns(24), 1121, Sheets("ワーク").Columns(111)) _
        + WorksheetFunction.SumIf(Sheets("ワーク").Columns(24), 1273, Sheets("ワーク").Columns(111)) _
        + WorksheetFunction.SumIf(Sheets("ワーク").Columns(24), 1166, Sheets("ワーク").Columns(111))

End Sub

