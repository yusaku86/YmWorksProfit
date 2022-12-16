Attribute VB_Name = "Module1"
'// YM”„Š|‹à—Œ`ì¬ƒ‚ƒWƒ…[ƒ‹
Option Explicit

'// ƒƒCƒ“ƒvƒƒV[ƒWƒƒ
Public Sub main()

    Call calculateTotalAmount

    Sheets("YM”„ã—Œ`").Activate

    Dim rs As New ADODB.Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM [ƒ[ƒN$]", connectDb(ThisWorkbook.FullName), adOpenStatic, adLockOptimistic

    rs.Sort = "¿‹æ”CˆÓCD ASC"

    Dim targetRow As Long: targetRow = 4

    '// Ø»²¸Ù—¿‚Ì‡Œv
    Dim totalRecyclingCharge
    
    Do Until rs.EOF
        If rs!¿‹æ”CˆÓCD = 0 Or rs!¿‹æ”CˆÓCD = 5013 Or rs!¿‹æ”CˆÓCD = 1121 Or rs!¿‹æ”CˆÓCD = 1273 Or rs!¿‹æ”CˆÓCD = 1166 Then
            GoTo DoContinue
        End If

        '// ‰ÛÅ‹àŠz“ü—Í
        Cells(targetRow, 1).Value = rs!”„ã‹æ•ªCD
        Cells(targetRow, 2).Value = rs!¿‹æ”CˆÓCD
        Cells(targetRow, 3).Value = rs!¿‹æ–¼1
        Cells(targetRow, 4).Value = rs!‰ÛÅ¬Œv
        Cells(targetRow, 5).Value = rs!Á”ïÅŒv
        Cells(targetRow, 14).Value = rs!‰ÛÅ¬Œv + rs!Á”ïÅŒv
        
        Cells(targetRow, 15).Value = rs!Ô—¼“o˜^”Ô†x‹Ç–¼ & rs!Ô—¼“o˜^”Ô†•ª—Ş & rs!Ô—¼“o˜^”Ô†‹L† & rs!Ô—¼“o˜^”Ô†”Ô† & " " & rs!”„ã‹æ•ª–¼Ì
        
        '// ©”…Ó‹àŠz‚ª0‰~‚æ‚è‘å‚«‚¢ê‡
        If rs!©”…Ó‹àŠz > 0 Then
            targetRow = targetRow + 1
            Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
            
            Cells(targetRow, 6).Value = rs!©”…Ó‹àŠz
            Cells(targetRow, 14).Value = rs!©”…Ó‹àŠz
        
            Cells(targetRow, 15).Value = rs!Ô—¼“o˜^”Ô†x‹Ç–¼ & rs!Ô—¼“o˜^”Ô†•ª—Ş & rs!Ô—¼“o˜^”Ô†‹L† & rs!Ô—¼“o˜^”Ô†”Ô† & " " & rs!”„ã‹æ•ª–¼Ì
        End If
        
        '// d—ÊÅ‹àŠz‚ª0‰~‚æ‚è‘å‚«‚¢ê‡
        If rs!d—ÊÅ‹àŠz > 0 Then
            targetRow = targetRow + 1
            Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
        
            Cells(targetRow, 7).Value = rs!d—ÊÅ‹àŠz
            Cells(targetRow, 14).Value = rs!d—ÊÅ‹àŠz
                    
            Cells(targetRow, 15).Value = rs!Ô—¼“o˜^”Ô†x‹Ç–¼ & rs!Ô—¼“o˜^”Ô†•ª—Ş & rs!Ô—¼“o˜^”Ô†‹L† & rs!Ô—¼“o˜^”Ô†”Ô† & " " & rs!”„ã‹æ•ª–¼Ì
        End If
        
        '/**
         '* ””ï—p‹àŠz‚ª0‰~‚æ‚è‘å‚«‚¢ê‡
         '* ¦ƒŠƒTƒCƒNƒ‹—¿‚ÍƒV[ƒg‚É•\¦‚µ‚È‚¢‚ªA‡Œv‹àŠz‚ª•K—v‚Ì‚½‚ß‡Œv‚·‚é
        '**/
        Dim i As Long
        
        For i = 1 To 5
            If rs.Fields("””ï—p‹àŠz" & i).Value <= 0 Then
                GoTo ForContinue
            End If
            
            '// ””ï—p‹àŠz1‚©‚ç””ï—p‹àŠz5‚Ü‚Å‹àŠz‚ğŠm”F‚µAí—Ş‚É‚æ‚Á‚Ä“ü—Í‚·‚éƒ}ƒX‚ğ•ÏX‚·‚é
            Select Case rs.Fields("””ï—p–¼Ì" & i).Value
            
                Case "ƒŠƒTƒCƒNƒ‹—¿"
                    totalRecyclingCharge = totalRecyclingCharge + rs.Fields("””ï—p‹àŠz" & i)
                    GoTo ForContinue
            
                Case "ŒŸ¸“o˜^ˆó†"
                    targetRow = targetRow + 1
                    Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
                    
                    Cells(targetRow, 8).Value = rs.Fields("””ï—p‹àŠz" & i).Value
                    Cells(targetRow, 15).Value = rs!Ô—¼“o˜^”Ô†x‹Ç–¼ & rs!Ô—¼“o˜^”Ô†•ª—Ş & rs!Ô—¼“o˜^”Ô†‹L† & rs!Ô—¼“o˜^”Ô†”Ô† & " " & rs!”„ã‹æ•ª–¼Ì
                
                Case "ÔŒŸˆó†‘ã"
                    targetRow = targetRow + 1
                    Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
                    
                    Cells(targetRow, 9).Value = rs.Fields("””ï—p‹àŠz" & i).Value
                    Cells(targetRow, 15).Value = rs!Ô—¼“o˜^”Ô†x‹Ç–¼ & rs!Ô—¼“o˜^”Ô†•ª—Ş & rs!Ô—¼“o˜^”Ô†‹L† & rs!Ô—¼“o˜^”Ô†”Ô† & " " & rs!”„ã‹æ•ª–¼Ì
                
                Case "—Õ‰^s‹–‰ÂØ"
                    targetRow = targetRow + 1
                    Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
                    
                    Cells(targetRow, 10).Value = rs.Fields("””ï—p‹àŠz" & i).Value
                    Cells(targetRow, 15).Value = rs!Ô—¼“o˜^”Ô†x‹Ç–¼ & rs!Ô—¼“o˜^”Ô†•ª—Ş & rs!Ô—¼“o˜^”Ô†‹L† & rs!Ô—¼“o˜^”Ô†”Ô† & " " & rs!”„ã‹æ•ª–¼Ì
                
                Case "“o˜^”Ô†•W"
                    targetRow = targetRow + 1
                    Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
                    
                    Cells(targetRow, 11).Value = rs.Fields("””ï—p‹àŠz" & i).Value
                    Cells(targetRow, 15).Value = rs!Ô—¼“o˜^”Ô†x‹Ç–¼ & rs!Ô—¼“o˜^”Ô†•ª—Ş & rs!Ô—¼“o˜^”Ô†‹L† & rs!Ô—¼“o˜^”Ô†”Ô† & " " & rs!”„ã‹æ•ª–¼Ì
                
                Case "Ô—¼”Ô†•W"
                    targetRow = targetRow + 1
                    Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
                    
                    Cells(targetRow, 12).Value = rs.Fields("””ï—p‹àŠz" & i).Value
                    Cells(targetRow, 15).Value = rs!Ô—¼“o˜^”Ô†x‹Ç–¼ & rs!Ô—¼“o˜^”Ô†•ª—Ş & rs!Ô—¼“o˜^”Ô†‹L† & rs!Ô—¼“o˜^”Ô†”Ô† & " " & rs!”„ã‹æ•ª–¼Ì
                
                Case "©“®ÔÅí•ÊŠ„"
                    targetRow = targetRow + 1
                    Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
                    
                    Cells(targetRow, 13).Value = rs.Fields("””ï—p‹àŠz" & i).Value
                    Cells(targetRow, 15).Value = rs!Ô—¼“o˜^”Ô†x‹Ç–¼ & rs!Ô—¼“o˜^”Ô†•ª—Ş & rs!Ô—¼“o˜^”Ô†‹L† & rs!Ô—¼“o˜^”Ô†”Ô† & " " & rs!”„ã‹æ•ª–¼Ì
            End Select
            
            Cells(targetRow, 14).Value = rs.Fields("””ï—p‹àŠz" & i).Value
    
ForContinue:
        Next
    
        targetRow = targetRow + 1
DoContinue:
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    '// ƒŠƒTƒCƒNƒ‹—¿‚Ì‡Œv“ü—Í
    Cells(2, 7).Value = totalRecyclingCharge
    
    '// ‡Œv‚ÌƒZƒ‹‚É®“ü—Í
    Cells(2, 5).Formula = "=SUM(N:N)"
    
End Sub

'// ‘ŒvAƒR[ƒh0(RŠİ‰^‘—)‚Ì‡ŒvAĞ“àŠÔ‡Œv‚ğŒvZ
Private Sub calculateTotalAmount()

    '// ‘Œv
    Cells(1, 5).Value = WorksheetFunction.Sum(Range(Sheets("ƒ[ƒN").Cells(2, 111), Sheets("ƒ[ƒN").Cells(Rows.Count, 111).End(xlUp)))
    
    '// ƒR[ƒh0
    Cells(1, 7).Value = WorksheetFunction.SumIf(Sheets("ƒ[ƒN").Columns(24), 0, Sheets("ƒ[ƒN").Columns(111))
    
    '// Ğ“àŠÔ
    Cells(1, 12).Value = _
        WorksheetFunction.SumIf(Sheets("ƒ[ƒN").Columns(24), 5013, Sheets("ƒ[ƒN").Columns(111)) _
        + WorksheetFunction.SumIf(Sheets("ƒ[ƒN").Columns(24), 1121, Sheets("ƒ[ƒN").Columns(111)) _
        + WorksheetFunction.SumIf(Sheets("ƒ[ƒN").Columns(24), 1273, Sheets("ƒ[ƒN").Columns(111)) _
        + WorksheetFunction.SumIf(Sheets("ƒ[ƒN").Columns(24), 1166, Sheets("ƒ[ƒN").Columns(111))

End Sub

