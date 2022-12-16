Attribute VB_Name = "functions"
Option Explicit

'// db接続
Public Function connectDb(ByVal dbBook As String) As ADODB.Connection

    Dim returnCon As New ADODB.Connection
    
    '// dbとして使用するファイルに接続
    With returnCon
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Properties("Extended Properties") = "Excel 12.0"
        .Open dbBook
    End With

    Set connectDb = returnCon

End Function
