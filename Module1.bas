Attribute VB_Name = "Module1"
'// YM|ΰ`μ¬W[
Option Explicit

'// CvV[W
Public Sub main()

    Call calculateTotalAmount

    Sheets("YMγ`").Activate

    Dim rs As New ADODB.Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM [[N$]", connectDb(ThisWorkbook.FullName), adOpenStatic, adLockOptimistic

    rs.Sort = "ΏζCΣCD ASC"

    Dim targetRow As Long: targetRow = 4

    '// Ψ»²ΈΩΏΜv
    Dim totalRecyclingCharge
    
    Do Until rs.EOF
        If rs!ΏζCΣCD = 0 Or rs!ΏζCΣCD = 5013 Or rs!ΏζCΣCD = 1121 Or rs!ΏζCΣCD = 1273 Or rs!ΏζCΣCD = 1166 Then
            GoTo DoContinue
        End If

        '// ΫΕΰzόΝ
        Cells(targetRow, 1).Value = rs!γζͺCD
        Cells(targetRow, 2).Value = rs!ΏζCΣCD
        Cells(targetRow, 3).Value = rs!ΏζΌ1
        Cells(targetRow, 4).Value = rs!ΫΕ¬v
        Cells(targetRow, 5).Value = rs!ΑοΕv
        Cells(targetRow, 14).Value = rs!ΫΕ¬v + rs!ΑοΕv
        
        Cells(targetRow, 15).Value = rs!ΤΌo^ΤxΗΌ & rs!ΤΌo^Τͺή & rs!ΤΌo^ΤL & rs!ΤΌo^ΤΤ & " " & rs!γζͺΌΜ
        
        '// ©Σΰzͺ0~ζθε«’κ
        If rs!©Σΰz > 0 Then
            targetRow = targetRow + 1
            Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
            
            Cells(targetRow, 6).Value = rs!©Σΰz
            Cells(targetRow, 14).Value = rs!©Σΰz
        
            Cells(targetRow, 15).Value = rs!ΤΌo^ΤxΗΌ & rs!ΤΌo^Τͺή & rs!ΤΌo^ΤL & rs!ΤΌo^ΤΤ & " " & rs!γζͺΌΜ
        End If
        
        '// dΚΕΰzͺ0~ζθε«’κ
        If rs!dΚΕΰz > 0 Then
            targetRow = targetRow + 1
            Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
        
            Cells(targetRow, 7).Value = rs!dΚΕΰz
            Cells(targetRow, 14).Value = rs!dΚΕΰz
                    
            Cells(targetRow, 15).Value = rs!ΤΌo^ΤxΗΌ & rs!ΤΌo^Τͺή & rs!ΤΌo^ΤL & rs!ΤΌo^ΤΤ & " " & rs!γζͺΌΜ
        End If
        
        '/**
         '* οpΰzͺ0~ζθε«’κ
         '* ¦TCNΏΝV[gΙ\¦΅Θ’ͺAvΰzͺKvΜ½ίv·ι
        '**/
        Dim i As Long
        
        For i = 1 To 5
            If rs.Fields("οpΰz" & i).Value <= 0 Then
                GoTo ForContinue
            End If
            
            '// οpΰz1©ηοpΰz5άΕΰzπmF΅AνήΙζΑΔόΝ·ι}XπΟX·ι
            Select Case rs.Fields("οpΌΜ" & i).Value
            
                Case "TCNΏ"
                    totalRecyclingCharge = totalRecyclingCharge + rs.Fields("οpΰz" & i)
                    GoTo ForContinue
            
                Case "Έo^σ"
                    targetRow = targetRow + 1
                    Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
                    
                    Cells(targetRow, 8).Value = rs.Fields("οpΰz" & i).Value
                    Cells(targetRow, 15).Value = rs!ΤΌo^ΤxΗΌ & rs!ΤΌo^Τͺή & rs!ΤΌo^ΤL & rs!ΤΌo^ΤΤ & " " & rs!γζͺΌΜ
                
                Case "Τσγ"
                    targetRow = targetRow + 1
                    Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
                    
                    Cells(targetRow, 9).Value = rs.Fields("οpΰz" & i).Value
                    Cells(targetRow, 15).Value = rs!ΤΌo^ΤxΗΌ & rs!ΤΌo^Τͺή & rs!ΤΌo^ΤL & rs!ΤΌo^ΤΤ & " " & rs!γζͺΌΜ
                
                Case "Υ^sΒΨ"
                    targetRow = targetRow + 1
                    Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
                    
                    Cells(targetRow, 10).Value = rs.Fields("οpΰz" & i).Value
                    Cells(targetRow, 15).Value = rs!ΤΌo^ΤxΗΌ & rs!ΤΌo^Τͺή & rs!ΤΌo^ΤL & rs!ΤΌo^ΤΤ & " " & rs!γζͺΌΜ
                
                Case "o^ΤW"
                    targetRow = targetRow + 1
                    Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
                    
                    Cells(targetRow, 11).Value = rs.Fields("οpΰz" & i).Value
                    Cells(targetRow, 15).Value = rs!ΤΌo^ΤxΗΌ & rs!ΤΌo^Τͺή & rs!ΤΌo^ΤL & rs!ΤΌo^ΤΤ & " " & rs!γζͺΌΜ
                
                Case "ΤΌΤW"
                    targetRow = targetRow + 1
                    Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
                    
                    Cells(targetRow, 12).Value = rs.Fields("οpΰz" & i).Value
                    Cells(targetRow, 15).Value = rs!ΤΌo^ΤxΗΌ & rs!ΤΌo^Τͺή & rs!ΤΌo^ΤL & rs!ΤΌo^ΤΤ & " " & rs!γζͺΌΜ
                
                Case "©?ΤΕνΚ"
                    targetRow = targetRow + 1
                    Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
                    
                    Cells(targetRow, 13).Value = rs.Fields("οpΰz" & i).Value
                    Cells(targetRow, 15).Value = rs!ΤΌo^ΤxΗΌ & rs!ΤΌo^Τͺή & rs!ΤΌo^ΤL & rs!ΤΌo^ΤΤ & " " & rs!γζͺΌΜ
            End Select
            
            Cells(targetRow, 14).Value = rs.Fields("οpΰz" & i).Value
    
ForContinue:
        Next
    
        targetRow = targetRow + 1
DoContinue:
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    '// TCNΏΜvόΝ
    Cells(2, 7).Value = totalRecyclingCharge
    
    '// vΜZΙ?όΝ
    Cells(2, 5).Formula = "=SUM(N:N)"
    
End Sub

'// vAR[h0(Rέ^)ΜvAΠΰΤvπvZ
Private Sub calculateTotalAmount()

    '// v
    Cells(1, 5).Value = WorksheetFunction.Sum(Range(Sheets("[N").Cells(2, 111), Sheets("[N").Cells(Rows.Count, 111).End(xlUp)))
    
    '// R[h0
    Cells(1, 7).Value = WorksheetFunction.SumIf(Sheets("[N").Columns(24), 0, Sheets("[N").Columns(111))
    
    '// ΠΰΤ
    Cells(1, 12).Value = _
        WorksheetFunction.SumIf(Sheets("[N").Columns(24), 5013, Sheets("[N").Columns(111)) _
        + WorksheetFunction.SumIf(Sheets("[N").Columns(24), 1121, Sheets("[N").Columns(111)) _
        + WorksheetFunction.SumIf(Sheets("[N").Columns(24), 1273, Sheets("[N").Columns(111)) _
        + WorksheetFunction.SumIf(Sheets("[N").Columns(24), 1166, Sheets("[N").Columns(111))

End Sub

