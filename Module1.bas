Attribute VB_Name = "Module1"
'// YM���|�����`�쐬���W���[��
Option Explicit

'// ���C���v���V�[�W��
Public Sub main()

    Call calculateTotalAmount

    Sheets("YM���㐗�`").Activate

    Dim rs As New ADODB.Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM [���[�N$]", connectDb(ThisWorkbook.FullName), adOpenStatic, adLockOptimistic

    rs.Sort = "������C��CD ASC"

    Dim targetRow As Long: targetRow = 4

    '// ػ��ٗ��̍��v
    Dim totalRecyclingCharge
    
    Do Until rs.EOF
        If rs!������C��CD = 0 Or rs!������C��CD = 5013 Or rs!������C��CD = 1121 Or rs!������C��CD = 1273 Or rs!������C��CD = 1166 Then
            GoTo DoContinue
        End If

        '// �ېŋ��z����
        Cells(targetRow, 1).Value = rs!����敪CD
        Cells(targetRow, 2).Value = rs!������C��CD
        Cells(targetRow, 3).Value = rs!�����於1
        Cells(targetRow, 4).Value = rs!�ېŏ��v
        Cells(targetRow, 5).Value = rs!����Ōv
        Cells(targetRow, 14).Value = rs!�ېŏ��v + rs!����Ōv
        
        Cells(targetRow, 15).Value = rs!�ԗ��o�^�ԍ��x�ǖ� & rs!�ԗ��o�^�ԍ����� & rs!�ԗ��o�^�ԍ��L�� & rs!�ԗ��o�^�ԍ��ԍ� & " " & rs!����敪����
        
        '// �����Ӌ��z��0�~���傫���ꍇ
        If rs!�����Ӌ��z > 0 Then
            targetRow = targetRow + 1
            Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
            
            Cells(targetRow, 6).Value = rs!�����Ӌ��z
            Cells(targetRow, 14).Value = rs!�����Ӌ��z
        
            Cells(targetRow, 15).Value = rs!�ԗ��o�^�ԍ��x�ǖ� & rs!�ԗ��o�^�ԍ����� & rs!�ԗ��o�^�ԍ��L�� & rs!�ԗ��o�^�ԍ��ԍ� & " " & rs!����敪����
        End If
        
        '// �d�ʐŋ��z��0�~���傫���ꍇ
        If rs!�d�ʐŋ��z > 0 Then
            targetRow = targetRow + 1
            Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
        
            Cells(targetRow, 7).Value = rs!�d�ʐŋ��z
            Cells(targetRow, 14).Value = rs!�d�ʐŋ��z
                    
            Cells(targetRow, 15).Value = rs!�ԗ��o�^�ԍ��x�ǖ� & rs!�ԗ��o�^�ԍ����� & rs!�ԗ��o�^�ԍ��L�� & rs!�ԗ��o�^�ԍ��ԍ� & " " & rs!����敪����
        End If
        
        '/**
         '* ����p���z��0�~���傫���ꍇ
         '* �����T�C�N�����̓V�[�g�ɕ\�����Ȃ����A���v���z���K�v�̂��ߍ��v����
        '**/
        Dim i As Long
        
        For i = 1 To 5
            If rs.Fields("����p���z" & i).Value <= 0 Then
                GoTo ForContinue
            End If
            
            '// ����p���z1���珔��p���z5�܂ŋ��z���m�F���A��ނɂ���ē��͂���}�X��ύX����
            Select Case rs.Fields("����p����" & i).Value
            
                Case "���T�C�N����"
                    totalRecyclingCharge = totalRecyclingCharge + rs.Fields("����p���z" & i)
                    GoTo ForContinue
            
                Case "�����o�^��"
                    targetRow = targetRow + 1
                    Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
                    
                    Cells(targetRow, 8).Value = rs.Fields("����p���z" & i).Value
                    Cells(targetRow, 15).Value = rs!�ԗ��o�^�ԍ��x�ǖ� & rs!�ԗ��o�^�ԍ����� & rs!�ԗ��o�^�ԍ��L�� & rs!�ԗ��o�^�ԍ��ԍ� & " " & rs!����敪����
                
                Case "�Ԍ��󎆑�"
                    targetRow = targetRow + 1
                    Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
                    
                    Cells(targetRow, 9).Value = rs.Fields("����p���z" & i).Value
                    Cells(targetRow, 15).Value = rs!�ԗ��o�^�ԍ��x�ǖ� & rs!�ԗ��o�^�ԍ����� & rs!�ԗ��o�^�ԍ��L�� & rs!�ԗ��o�^�ԍ��ԍ� & " " & rs!����敪����
                
                Case "�Վ��^�s����"
                    targetRow = targetRow + 1
                    Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
                    
                    Cells(targetRow, 10).Value = rs.Fields("����p���z" & i).Value
                    Cells(targetRow, 15).Value = rs!�ԗ��o�^�ԍ��x�ǖ� & rs!�ԗ��o�^�ԍ����� & rs!�ԗ��o�^�ԍ��L�� & rs!�ԗ��o�^�ԍ��ԍ� & " " & rs!����敪����
                
                Case "�o�^�ԍ��W"
                    targetRow = targetRow + 1
                    Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
                    
                    Cells(targetRow, 11).Value = rs.Fields("����p���z" & i).Value
                    Cells(targetRow, 15).Value = rs!�ԗ��o�^�ԍ��x�ǖ� & rs!�ԗ��o�^�ԍ����� & rs!�ԗ��o�^�ԍ��L�� & rs!�ԗ��o�^�ԍ��ԍ� & " " & rs!����敪����
                
                Case "�ԗ��ԍ��W"
                    targetRow = targetRow + 1
                    Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
                    
                    Cells(targetRow, 12).Value = rs.Fields("����p���z" & i).Value
                    Cells(targetRow, 15).Value = rs!�ԗ��o�^�ԍ��x�ǖ� & rs!�ԗ��o�^�ԍ����� & rs!�ԗ��o�^�ԍ��L�� & rs!�ԗ��o�^�ԍ��ԍ� & " " & rs!����敪����
                
                Case "�����ԐŎ�ʊ�"
                    targetRow = targetRow + 1
                    Range(Cells(targetRow, 1), Cells(targetRow, 3)).Value = Range(Cells(targetRow - 1, 1), Cells(targetRow - 1, 3)).Value
                    
                    Cells(targetRow, 13).Value = rs.Fields("����p���z" & i).Value
                    Cells(targetRow, 15).Value = rs!�ԗ��o�^�ԍ��x�ǖ� & rs!�ԗ��o�^�ԍ����� & rs!�ԗ��o�^�ԍ��L�� & rs!�ԗ��o�^�ԍ��ԍ� & " " & rs!����敪����
            End Select
            
            Cells(targetRow, 14).Value = rs.Fields("����p���z" & i).Value
    
ForContinue:
        Next
    
        targetRow = targetRow + 1
DoContinue:
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    '// ���T�C�N�����̍��v����
    Cells(2, 7).Value = totalRecyclingCharge
    
    '// ���v�̃Z���Ɏ�����
    Cells(2, 5).Formula = "=SUM(N:N)"
    
End Sub

'// ���v�A�R�[�h0(�R�݉^��)�̍��v�A�Г��ԍ��v���v�Z
Private Sub calculateTotalAmount()

    '// ���v
    Cells(1, 5).Value = WorksheetFunction.Sum(Range(Sheets("���[�N").Cells(2, 111), Sheets("���[�N").Cells(Rows.Count, 111).End(xlUp)))
    
    '// �R�[�h0
    Cells(1, 7).Value = WorksheetFunction.SumIf(Sheets("���[�N").Columns(24), 0, Sheets("���[�N").Columns(111))
    
    '// �Г���
    Cells(1, 12).Value = _
        WorksheetFunction.SumIf(Sheets("���[�N").Columns(24), 5013, Sheets("���[�N").Columns(111)) _
        + WorksheetFunction.SumIf(Sheets("���[�N").Columns(24), 1121, Sheets("���[�N").Columns(111)) _
        + WorksheetFunction.SumIf(Sheets("���[�N").Columns(24), 1273, Sheets("���[�N").Columns(111)) _
        + WorksheetFunction.SumIf(Sheets("���[�N").Columns(24), 1166, Sheets("���[�N").Columns(111))

End Sub

