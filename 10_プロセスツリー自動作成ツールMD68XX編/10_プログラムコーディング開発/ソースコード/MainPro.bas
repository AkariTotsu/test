Attribute VB_Name = "MainPro"
'**********************************************************************************************
' 図面番号　                ：FSP01100800003
' 名称　                    ：プロセスツリー自動作成ツールMD68XX編図面
' ソフトウェア管理番号ＩＤ　：FSS0101037
' ソフトウェア名　          ：プロセスツリー自動作成ツールMD68XX編
' モジュール名　            ：MainPro
' 機能概要　　　            ：プロジェクト管理時、プロセスツリーの作成と運用管理を自動化する。
'
' 改訂履歴　　　            ：2019/07/24 新規 T.Takayasu
'
' Copyright(C) 2019 Sanken Electric Co., Ltd. All Rights Reserved.
'**********************************************************************************************

Option Explicit

'====================================================================================================
' Main CreateNewFolder
'   \\Hka0-fspc\c\Users\GITCLONE\Documents\GIT_CLONE
'   作業フォルダ ---> 無
'====================================================================================================
' CONTENTS (プロセスツリー自動作成ツール機能安全課編)
'   CreateNewFolder
'   DeleteEmptyMatrix
'   DeleteEmptyRowsColumns
'   CreateParentFolder
'   CheckForProhibitedCharacters
'   FillInTheBlanks
'   LastRow
'   LastColumn
'   LastRow_MAX
'   LastColumn_MAX
'=====================================================================================================
' DEPLOYMENT
'   (1)設置場所：任意
'   (2)A1セルにプロセスツリーの作成場所を指定（\\Hka0-fspc\c\Users\GITCLONE\Documents\GIT_CLONE）
'   (3)A2セルにプロセスツリーの最大階層数を記入
'   (4)【機能安全プロセスツリー作成】ボタンを選択・実行
'   (5)機能安全課プロセスツリーの変更時は必ず「プロセスツリー自動作成ツール機能安全課編」を使用すること
'======================================================================================================
'
'-------------------------------------------------------------------------------
' Dim Table
'-------------------------------------------------------------------------------

Public FolderCounter As Long      ' フォルダ数のカウンタ
Public InterruptionFlg As Boolean   ' 中断処理のフラグ


'-------------------------------------------------------------------------------
' CreateNewFolder
'-------------------------------------------------------------------------------

Sub CreateNewFolder()
' ---- ★新規フォルダを作成するプログラム★（シートはいじらず，メモリ上で実現） ----

    Dim lngR As Long
    Dim lngC As Long
    Dim vntTmp As Variant
    Dim strBuf As String
    Dim sngT As Single
    Dim lngMsg As Long

    On Error Resume Next                                                            ' Err.Numberと組み合わせて使う

    ' ---- DeleteEmptyRowsColumns ----
    Call DeleteEmptyRowsColumns
    
    InterruptionFlg = Not InterruptionFlg                                           ' ここでTrueとなる
    If InterruptionFlg = True Then
        ' ---- 親フォルダの設定 ----
        If Cells(1, 1) <> "" Then
            lngMsg = MsgBox("A1に記入されているフォルダにツリーを作成します。", vbQuestion + vbYesNoCancel) ' ★
            If lngMsg = vbCancel Then GoTo Exit_Handler                                     ' ★lngMsgがCancelの時
        End If
        
        If Cells(1, 1) = "" Or lngMsg = vbNo Then                                   ' A1が空か，★lngMsgがNoのとき
            MsgBox "保存フォルダが指定されていません。"
            GoTo Exit_Handler
        End If
        
        ' ---- 禁止文字のチェック ----
        If CheckForProhibitedCharacters.Address <> "$A$1" Then                      ' なければ，$A$1が返る
            CheckForProhibitedCharacters.Select
            MsgBox "次のセルに使用禁止の文字が含まれています。" & vbLf & CheckForProhibitedCharacters.Address, vbCritical
            GoTo Exit_Handler
        End If
        
    End If

    ' ---- FillInTheBlanks ----
    vntTmp = Cells(1, 1).Resize(LastRow_MAX, LastColumn_MAX)                        ' 配列として格納
    vntTmp = FillInTheBlanks(vntTmp)                                                ' 変数上で
    
    ' ---- ★CreateParentFolder★ ----
    FolderCounter = 0
    Call CreateParentFolder(Cells(1, 1))
    
    sngT = Timer
    ' ---- ★新規フォルダの作成★ ----
    For lngR = 2 To LastRow_MAX
        strBuf = vntTmp(1, 1)

        ' ---- 作成 ----
        For lngC = 2 To LastColumn(lngR)
            strBuf = strBuf & "\" & vntTmp(lngR, lngC)                             ' 文字列をつないで
            
            Err.Number = 0
            MkDir (strBuf) ' フォルダの作成
            If Err.Number = 0 Then FolderCounter = FolderCounter + 1                ' 作成できたら（エラーじゃなければ）カウンタを＋１
            Application.StatusBar = "作成フォルダ数：　" & FolderCounter

            DoEvents

            If InterruptionFlg = False Then
                MsgBox "中断しました。", vbInformation
                GoTo Exit_Handler
            End If
        Next lngC
    Next lngR

    If FolderCounter > 0 Then
        MsgBox "完了　" & Timer - sngT & " sec", vbInformation
    Else
        MsgBox "フォルダが作成できませんでした。" & vbLf & "次のような原因が考えられます。" & vbLf & vbLf & "・ドライブがない。" & vbLf & "・同じフォルダが既にある。　など", vbCritical
    End If
    
    ActiveSheet.Hyperlinks.Delete                                                   ' ハイパーリンクを解除する

Exit_Handler:
    Application.StatusBar = False
    InterruptionFlg = False                                                         ' ここで元に戻す(Public変数)

End Sub

'-------------------------------------------------------------------------------
' DeleteEmptyMatrix
'-------------------------------------------------------------------------------

Sub DeleteEmptyMatrix()

    Call DeleteEmptyRowsColumns
    
    MsgBox "完了", vbInformation

End Sub

'-------------------------------------------------------------------------------
' DeleteEmptyRowsColumns
'-------------------------------------------------------------------------------

Sub DeleteEmptyRowsColumns()

    Dim lngR As Long
    Dim lngC As Long

    ' ---- 空行の削除 ----
    For lngR = LastRow_MAX To 2 Step -1
        If Application.WorksheetFunction.CountBlank(Rows(lngR)) = Columns.Count Then
            Rows(lngR).Delete
        End If
    Next lngR

    ' ---- 空列の削除 ----
    For lngC = LastColumn_MAX To 2 Step -1
        If Application.WorksheetFunction.CountBlank(Columns(lngC)) = Rows.Count Then
            Columns(lngC).Delete
'            MsgBox "lngC=" & lngC
        End If
    Next lngC

End Sub

'-------------------------------------------------------------------------------
' CreateParentFolder
'-------------------------------------------------------------------------------

Sub CreateParentFolder(strBuf As String)
' ---- CreateParentFolder ----

    Dim vntTmp As Variant
    Dim strT As String
    Dim lngN As Long
    Dim strNetwork As String
    Dim strFolder As String
    Dim objFSO As Object
    Dim lngMsg As Long
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    strFolder = strBuf
    Err.Number = 0
    
    On Error Resume Next
    
    ' ---- ファイルシステムオブジェクトを使って，親フォルダの有無を調べる ----Dir(strT, vbDirectory)では，ドライブが空だと検出できないので
    If objFSO.FolderExists(strFolder) = False Then
        lngMsg = MsgBox("親フォルダが存在しません。" & vbCr & "新規で作って処理を続けますか？", vbYesNo + vbQuestion)
    End If

    ' ---- CreateParentFolder処理をする ----
    If lngMsg = vbYes Then                                                  ' ＜新規作成＞
    
        ' ---- MkDirで作るための下準備 ----
        If Left(strBuf, 2) = "\\" Then
            strBuf = Right(strBuf, Len(strBuf) - 2)                         ' ネットワーク対応（先頭の\\を取り除く）
            strNetwork = "\\"
        End If
        
        vntTmp = Split(strBuf, "\")
        strT = strNetwork & vntTmp(0)                                       ' ネットワーク対応（先頭に\\を付ける）
        
        ' ---- 作成処理 ----
        For lngN = LBound(vntTmp) + 1 To UBound(vntTmp)
            strT = strT & "\" & vntTmp(lngN)
    
            Err.Number = 0
            MkDir (strT)                                                    ' フォルダの作成
            If Err.Number = 0 Then FolderCounter = FolderCounter + 1        ' 作成できたら（エラーじゃなければ）カウンタを＋１
    
            Application.StatusBar = "作成フォルダ数：　" & FolderCounter
        Next lngN

        
    ElseIf lngMsg = vbNo Then                                               ' ＜中断＞
        InterruptionFlg = False                                             ' ここで元に戻す(Public変数)
        Application.StatusBar = False

        End
    End If
    
End Sub

'-------------------------------------------------------------------------------
' CheckForProhibitedCharacters
'-------------------------------------------------------------------------------

Function CheckForProhibitedCharacters() As Range

' ---- 禁止文字のチェック ----

    Dim lngN As Long
    Dim vntBan As Variant
    Dim objRg As Range                                                      ' 時間短縮のため
    
    vntBan = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    
    For Each objRg In Cells(1, 2).Resize(LastRow_MAX, LastColumn_MAX - 1)
        For lngN = LBound(vntBan) To UBound(vntBan)
            If InStr(objRg.Value, vntBan(lngN)) > 0 Then                    ' 禁止文字があれば位置を返すので
                Set CheckForProhibitedCharacters = objRg
                Exit Function
            End If
        Next lngN
    Next objRg

    Set CheckForProhibitedCharacters = Cells(1, 1)                          ' 禁止文字がなければ，$A$1を返す

End Function

'-------------------------------------------------------------------------------
' FillInTheBlanks
'-------------------------------------------------------------------------------

Function FillInTheBlanks(vntTmp As Variant) As Variant

' ---- セルで言うところの，FillInTheBlanks ----

    Dim lngR As Long
    Dim lngC As Long
    
    ' ---- 2列目までの処理 ----
    For lngC = 1 To 2
        For lngR = 2 To LastRow_MAX
            If vntTmp(lngR, lngC) = "" Then
                vntTmp(lngR, lngC) = vntTmp(lngR - 1, lngC)
            End If
        Next lngR
    Next lngC
    
    ' ---- 3列目以降の処理----
    For lngC = 3 To LastColumn_MAX
        For lngR = 2 To LastRow_MAX
            If vntTmp(lngR, lngC) = "" Then
                If vntTmp(lngR, lngC - 1) = vntTmp(lngR - 1, lngC - 1) Then
                    vntTmp(lngR, lngC) = vntTmp(lngR - 1, lngC)
                End If
            End If
        Next lngR
    Next lngC
    
    FillInTheBlanks = vntTmp

End Function

'-------------------------------------------------------------------------------
' LastRow
'-------------------------------------------------------------------------------

Function LastRow(last_column As Long) As Long
' ---- その列のLastRowを求める ----

    Dim lngR As Long
    
    lngR = Cells(Rows.Count, last_column).End(xlUp).Row                         ' 行末から探す
    If Cells(1, last_column) = "" And lngR = 1 Then lngR = 0

    LastRow = lngR

End Function

'-------------------------------------------------------------------------------
' LastColumn
'-------------------------------------------------------------------------------

Function LastColumn(last_row As Long) As Long
' ---- その行のLastColumnを求める ----

    Dim lngC As Long
    
    lngC = Cells(last_row, Columns.Count).End(xlToLeft).Column                  ' 列末から探す
    If Cells(last_row, 1) = "" And lngC = 1 Then lngC = 0

    LastColumn = lngC

End Function

'-------------------------------------------------------------------------------
' LastRow_MAX
'-------------------------------------------------------------------------------

Function LastRow_MAX() As Long
' ---- データの入っている最大LastRowを求める ----

    LastRow_MAX = ActiveSheet.UsedRange.Item(ActiveSheet.UsedRange.Count).Row   ' 使用している範囲での最後のセルの行を求める
    
End Function

'-------------------------------------------------------------------------------
' LastColumn_MAX
'-------------------------------------------------------------------------------

Function LastColumn_MAX() As Long
' ---- データの入っている最大LastColumnを求める ----

    LastColumn_MAX = Range("A2").Value                                          ' 最大階層をA2の値に固定
    LastColumn_MAX = LastColumn_MAX + 1

End Function
