Attribute VB_Name = "programParts"
Option Explicit

'**********　SC番号取得　**********

'   引数    なし

'   戻り値
'       integer   SC番号

'****************************************

Function SCNumber() As Variant

    Application.StatusBar = "SCNumber Start !!"
    Debug.Print "SCNumber Start !!"

    Dim SCNumRow As Integer
    Dim SCNumCol As Integer

'SCナンバー取得
    
    SCNumRow = 1
    SCNumCol = 1
    
    'エラーを無視
    On Error Resume Next
    'SC番号の位置を検索
    Do Until Cells(SCNumRow, SCNumCol) Like "SC番号"
        SCNumRow = SCNumRow + 1
        If Err.Number = 6 Then
            SCNumCol = SCNumCol + 1
            SCNumRow = 1
            Err.Clear
        End If
        If Err.Number <> 0 Or SCNumCol > 10 Then
            MsgBox "エラーが発生した為、終了します。" + vbCrLf + "[SC番号取得エラー]"
            End
        End If
    Loop
    'エラー無視を解除
    On Error GoTo 0

    SCNumber = Str(Cells(SCNumRow, SCNumCol + 1))
    
    Application.StatusBar = "SCNumber Finish !!"
    Debug.Print "SCNumber Finish !!"
    
End Function

'**********　取得済みの日付取得　**********
'
'   引数
'       なし
'
'   戻り値
'       String()        記載済み日付(配列)

'******************************************

Function retrievedDate() As String()

'【変数】
    Dim datePointRow As Integer
    Dim endPoint As Long
    Dim dates() As String
    Dim i As Integer
    
    '日付の位置取得
    datePointRow = 1
    Do Until Cells(datePointRow, 1) = "日付"
        datePointRow = datePointRow + 1
    Loop
    
    '最終行取得
    endPoint = Cells(Rows.Count, 1).End(xlUp).Row
    
    '配列再定義
    If endPoint <> datePointRow Then
        ReDim dates(endPoint - datePointRow - 1)
        
        '配列に日付を代入
        i = 1
        Do Until Cells(datePointRow + i, 1) = ""
            dates(i - 1) = Format(Cells(datePointRow + i, 1), "yyyymmdd")
            i = i + 1
        Loop
     
        '取得した日付を戻り値に設定
        retrievedDate = dates
    
     End If
     
End Function


'**********　CSVファイル読込　**********

'   引数
'       endDate(string)     入力済み日付
'       SCnum(variant)      検索コード
'       titlechar(variant)  タイトル

'   戻り値
'        String()    全ファイルの全文字列
        
'****************************************

Sub CSVFileRead(ByRef endDate() As String, ByVal SCNum As Variant)
    
'【変数】
    Dim i As Integer 'カウンター
    Dim myFName As String 'ファイル名
    Dim myFPath As String 'ファイルパス
    Dim myFNo As Integer '開くファイルの番号
    Dim myBuf As String '取得した文字列
    Dim sample() As String
    Dim chkDate As Variant
    
    'フォルダ内にある全ファイルパスを取得
    myFName = Dir(ThisWorkbook.Path & Application.PathSeparator + "*.csv")
    
    '全ファイルの文字列を取得
    i = 0
    Do Until myFName = ""
        '入力済みの日付をスキップ
'        If UBound(endDate) <> 0 Then
'        For Each chkDate In endDate
'            If myFName Like "*" & chkDate & "*" Then
'                GoTo nxt1
'            End If
'        Next
'        End If
        
        'ファイルパスを作成
        myFPath = ThisWorkbook.Path & Application.PathSeparator & myFName
        'ファイル番号を付与
        myFNo = FreeFile
        'ファイルを開く
        Open myFPath For Input As #myFNo
        '開いたファイルを変数へ代入
        myBuf = StrConv(InputB(LOF(myFNo), #myFNo), vbUnicode)
        'ファイルを閉じる
        Close #myFNo
        Call charShaping(myBuf, SCNum)
        i = i + 1
        
nxt1:
        '次のファイルへ
        myFName = Dir()
    Loop
        
End Sub


'**********　一行ずつ変数へ代入　**********



'****************************************

Sub charShaping(ByRef myBuf As String, ByVal SCNum As Variant)

'【定数】
    Const fRETURNCODE As String = vbCrLf '改行文字
    Const fDELIMITER As String = ","  '区切り文字

'【変数】
    Dim myRowArray() As String
    Dim title() As String
    Dim chkSCNum As Variant
    Dim Provisional As Variant
    Dim i As Integer
    
    myRowArray = Split(myBuf, fRETURNCODE)
    title = Split(myRowArray(0), fDELIMITER)

    'SCの位置取得
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


'**********　セルへ入力　**********



'**********************************

Sub assignToCell(ByVal targetLine As Variant)

'【変数】
    Dim titleChar() As String
    Dim titlePointRow As Integer
    Dim endPointRow As Long

    '入力箇所取得
    titleChar = getTitle
    
    titlePointRow = 1
    Do Until Cells(titlePointRow, 1) = "日付"
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


'**********　メインエクセルのタイトル文字取得　**********

'   引数    なし

'   戻り値
'        variant     タイトル文字(配列)

'********************************************************

Function getTitle() As String()

'【変数】
    Dim titlePointRow As Integer '横位置
    Dim titlePointCol As Integer '縦位置
    Dim title() As String 'タイトル文字を入れる配列
    Dim i As Integer
    
'縦位置取得
    titlePointRow = 1
    Do Until Cells(titlePointRow, 1) Like "日付"
        titlePointRow = titlePointRow + 1
    Loop
    
'横位置取得
    titlePointCol = 1
    Do Until Cells(titlePointRow, titlePointCol) = ""
        titlePointCol = titlePointCol + 1
    Loop
    titlePointCol = titlePointCol - 1
    
 '配列定義
    ReDim title(titlePointCol - 1)
    
'配列にタイトル文字を代入
    i = 1
    Do Until Cells(titlePointRow, i) = ""
        title(i - 1) = Cells(titlePointRow, i)
        i = i + 1
    Loop
    
'戻り値設定
    getTitle = title
    
End Function
