Attribute VB_Name = "mainProgram"
Option Explicit


'**********�@���C���v���O�����@**********
Sub getCharactersInCsvFile()

'�y�ϐ��z
    Dim SCNum As Variant
    Dim myBuf As String
    Dim endDate() As String
    
'SC�ԍ��擾
    SCNum = Replace(SCNumber, " ", "")
    
'���͍ς݂̓��t���擾
    endDate = retrievedDate
    
'�S�t�@�C���̑S��������擾
    Call CSVFileRead(endDate, SCNum)
    
    

    
End Sub


