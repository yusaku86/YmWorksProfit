Attribute VB_Name = "functions"
Option Explicit

'// db�ڑ�
Public Function connectDb(ByVal dbBook As String) As ADODB.Connection

    Dim returnCon As New ADODB.Connection
    
    '// db�Ƃ��Ďg�p����t�@�C���ɐڑ�
    With returnCon
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Properties("Extended Properties") = "Excel 12.0"
        .Open dbBook
    End With

    Set connectDb = returnCon

End Function
