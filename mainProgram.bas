Attribute VB_Name = "mainProgram"
Option Explicit


'**********　メインプログラム　**********
Sub getCharactersInCsvFile()

'【変数】
    Dim SCNum As Variant
    Dim myBuf As String
    Dim endDate() As String
    
'SC番号取得
    SCNum = Replace(SCNumber, " ", "")
    
'入力済みの日付を取得
    endDate = retrievedDate
    
'全ファイルの全文字列を取得
    Call CSVFileRead(endDate, SCNum)
    
    

    
End Sub


