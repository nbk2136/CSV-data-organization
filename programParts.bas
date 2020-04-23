Attribute VB_Name = "programParts"
Option Explicit

'**********�@SC�ԍ��擾�@**********

'   ����    �Ȃ�

'   �߂�l
'       integer   SC�ԍ�

'****************************************

Function SCNumber() As Variant

    Application.StatusBar = "SCNumber Start !!"
    Debug.Print "SCNumber Start !!"

    Dim SCNumRow As Integer
    Dim SCNumCol As Integer

'SC�i���o�[�擾
    
    SCNumRow = 1
    SCNumCol = 1
    
    '�G���[�𖳎�
    On Error Resume Next
    'SC�ԍ��̈ʒu������
    Do Until Cells(SCNumRow, SCNumCol) Like "SC�ԍ�"
        SCNumRow = SCNumRow + 1
        If Err.Number = 6 Then
            SCNumCol = SCNumCol + 1
            SCNumRow = 1
            Err.Clear
        End If
        If Err.Number <> 0 Or SCNumCol > 10 Then
            MsgBox "�G���[�����������ׁA�I�����܂��B" + vbCrLf + "[SC�ԍ��擾�G���[]"
            End
        End If
    Loop
    '�G���[����������
    On Error GoTo 0

    SCNumber = Str(Cells(SCNumRow, SCNumCol + 1))
    
    Application.StatusBar = "SCNumber Finish !!"
    Debug.Print "SCNumber Finish !!"
    
End Function

'**********�@�擾�ς݂̓��t�擾�@**********
'
'   ����
'       �Ȃ�
'
'   �߂�l
'       String()        �L�ڍςݓ��t(�z��)

'******************************************

Function retrievedDate() As String()

'�y�ϐ��z
    Dim datePointRow As Integer
    Dim endPoint As Long
    Dim dates() As String
    Dim i As Integer
    
    '���t�̈ʒu�擾
    datePointRow = 1
    Do Until Cells(datePointRow, 1) = "���t"
        datePointRow = datePointRow + 1
    Loop
    
    '�ŏI�s�擾
    endPoint = Cells(Rows.Count, 1).End(xlUp).Row
    
    '�z��Ē�`
    If endPoint <> datePointRow Then
        ReDim dates(endPoint - datePointRow - 1)
        
        '�z��ɓ��t����
        i = 1
        Do Until Cells(datePointRow + i, 1) = ""
            dates(i - 1) = Format(Cells(datePointRow + i, 1), "yyyymmdd")
            i = i + 1
        Loop
     
        '�擾�������t��߂�l�ɐݒ�
        retrievedDate = dates
    
     End If
     
End Function


'**********�@CSV�t�@�C���Ǎ��@**********

'   ����
'       endDate(string)     ���͍ςݓ��t
'       SCnum(variant)      �����R�[�h
'       titlechar(variant)  �^�C�g��

'   �߂�l
'        String()    �S�t�@�C���̑S������
        
'****************************************

Sub CSVFileRead(ByRef endDate() As String, ByVal SCNum As Variant)
    
'�y�ϐ��z
    Dim i As Integer '�J�E���^�[
    Dim myFName As String '�t�@�C����
    Dim myFPath As String '�t�@�C���p�X
    Dim myFNo As Integer '�J���t�@�C���̔ԍ�
    Dim myBuf As String '�擾����������
    Dim sample() As String
    Dim chkDate As Variant
    
    '�t�H���_���ɂ���S�t�@�C���p�X���擾
    myFName = Dir(ThisWorkbook.Path & Application.PathSeparator + "*.csv")
    
    '�S�t�@�C���̕�������擾
    i = 0
    Do Until myFName = ""
        '���͍ς݂̓��t���X�L�b�v
'        If UBound(endDate) <> 0 Then
'        For Each chkDate In endDate
'            If myFName Like "*" & chkDate & "*" Then
'                GoTo nxt1
'            End If
'        Next
'        End If
        
        '�t�@�C���p�X���쐬
        myFPath = ThisWorkbook.Path & Application.PathSeparator & myFName
        '�t�@�C���ԍ���t�^
        myFNo = FreeFile
        '�t�@�C�����J��
        Open myFPath For Input As #myFNo
        '�J�����t�@�C����ϐ��֑��
        myBuf = StrConv(InputB(LOF(myFNo), #myFNo), vbUnicode)
        '�t�@�C�������
        Close #myFNo
        Call charShaping(myBuf, SCNum)
        i = i + 1
        
nxt1:
        '���̃t�@�C����
        myFName = Dir()
    Loop
        
End Sub


'**********�@��s���ϐ��֑���@**********



'****************************************

Sub charShaping(ByRef myBuf As String, ByVal SCNum As Variant)

'�y�萔�z
    Const fRETURNCODE As String = vbCrLf '���s����
    Const fDELIMITER As String = ","  '��؂蕶��

'�y�ϐ��z
    Dim myRowArray() As String
    Dim title() As String
    Dim chkSCNum As Variant
    Dim Provisional As Variant
    Dim i As Integer
    
    myRowArray = Split(myBuf, fRETURNCODE)
    title = Split(myRowArray(0), fDELIMITER)

    'SC�̈ʒu�擾
     i = 0
    For Each Provisional In title
        If Provisional Like "*" & "SC" & "*" Then
            Exit For
        End If
        i = i + 1
    Next
    
    For Each chkSCNum In myRowArray
        chkSCNum = Split(chkSCNum, fDELIMITER)
        If chkSCNum(i) Like "*" & SCNum & "*" Then
            Call assignToCell(chkSCNum)
        End If
    Next
    
    myRowArray = myRowArray
End Sub


'**********�@�Z���֓��́@**********



'**********************************

Sub assignToCell(ByVal targetLine As Variant)

'�y�ϐ��z
    Dim titleChar() As String
    Dim titlePointRow As Integer
    Dim endPointRow As Long

    '���͉ӏ��擾
    titleChar = getTitle
    
    titlePointRow = 1
    Do Until Cells(titlePointRow, 1) = "���t"
        titlePointRow = titlePointRow + 1
    Loop
    titlePointRow = titlePointRow + 1
    
    endPointRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Cells(endPointRow, 1) = targetLine(4)
    Cells(endPointRow, 2) = targetLine(9)
    Cells(endPointRow, 3) = targetLine(10)
    Cells(endPointRow, 4) = targetLine(11)
    Cells(endPointRow, 5) = targetLine(5)
    Cells(endPointRow, 6) = targetLine(12)
    

End Sub


'**********�@���C���G�N�Z���̃^�C�g�������擾�@**********

'   ����    �Ȃ�

'   �߂�l
'        variant     �^�C�g������(�z��)

'********************************************************

Function getTitle() As String()

'�y�ϐ��z
    Dim titlePointRow As Integer '���ʒu
    Dim titlePointCol As Integer '�c�ʒu
    Dim title() As String '�^�C�g������������z��
    Dim i As Integer
    
'�c�ʒu�擾
    titlePointRow = 1
    Do Until Cells(titlePointRow, 1) Like "���t"
        titlePointRow = titlePointRow + 1
    Loop
    
'���ʒu�擾
    titlePointCol = 1
    Do Until Cells(titlePointRow, titlePointCol) = ""
        titlePointCol = titlePointCol + 1
    Loop
    titlePointCol = titlePointCol - 1
    
 '�z���`
    ReDim title(titlePointCol - 1)
    
'�z��Ƀ^�C�g����������
    i = 1
    Do Until Cells(titlePointRow, i) = ""
        title(i - 1) = Cells(titlePointRow, i)
        i = i + 1
    Loop
    
'�߂�l�ݒ�
    getTitle = title
    
End Function
