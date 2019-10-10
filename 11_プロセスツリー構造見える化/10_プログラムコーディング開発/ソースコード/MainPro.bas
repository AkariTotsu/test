Attribute VB_Name = "MainPro"
'**********************************************************************************************
' 図面番号　                ：FSP01110800001
' 名称　                    ：プロセスツリー構造見える化
' ソフトウェア管理番号ＩＤ　：FSS0101038
' ソフトウェア名　          ：プロセスツリー構造見える化
' モジュール名　            ：MainPro
' 機能概要　　　            ：プロセスツリーの構造をエクセルシート上に自動作成し見える化する。
'
' 改訂履歴　　　            ：2019/08/16 新規 T.Takayasu
'
' Copyright(C) 2019 Sanken Electric Co., Ltd. All Rights Reserved.
'**********************************************************************************************

Option Explicit

'====================================================================================================
' Main CreateNewFolder
'   \\Hka0-fspc\c\Users\GITCLONE\Documents\GIT_CLONE
'   作業フォルダ ---> 無
'====================================================================================================
' CONTENTS (プロセスツリー構造見える化)
'   ExportFolderStructure
'   ExportSubfolders
'   DeleteIfSame
'   DeletingEmptyRowsColumns
'   FindLastLine_MAX
'   FindLastColumn_MAX
'=====================================================================================================
' DEPLOYMENT
'   (1)設置場所：任意
'   (2)B3セルにプロセスツリーの初期フォルダを指定（C:\Users\takayasu-toshiyuki\Documents\プロセスツリー作成テスト用\）
'   (3)【プロセスツリー構造見える化】ボタンを選択・実行
'======================================================================================================
'
'-------------------------------------------------------------------------------
' Dim Table
'-------------------------------------------------------------------------------

Public CountFolder As Long                                  'フォルダ数のカウンタ
Public InterruptionFlg As Boolean                           '中断処理のフラグ
Dim Class As Long                                           '階層のカウンタ
Dim Limitation As Variant                                   '階層の深さ制限

'-------------------------------------------------------------------------------
' フォルダの構造を書き出すプログラム
'-------------------------------------------------------------------------------
Sub ExportFolderStructure()

Dim WshShell
Dim dataFolder          As Variant
Dim lngR As Long
Dim strFD As String
Dim strFDName As String
Dim strConnectionCharacter As String
Dim sglT As Single
Dim lngRetsu As Long

    InterruptionFlg = Not InterruptionFlg                                               '初期値は，ここでTrueになる
    
    Sheets("INDEX").Select

    dataFolder = Range("B3").Value                                                      '初期フォルダ

    If InterruptionFlg = True Then
        '---- 処理するフォルダを選択する ----
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "プロセスツリー構造を読み込むフォルダを選択して下さい。（実行中に，もう一度行うと中断します。）"
            .InitialFileName = dataFolder                                               '初期フォルダを指定する
            If .Show = False Then GoTo Exit_Handler                                     'ダイアログボックスを表示
            strFD = .SelectedItems(1)                                                   'フォルダ名を格納
        End With
        
    strFDName = Right(strFD, InStr(StrReverse(strFD), "\") - 1)                         '最下層のフォルダ名を取得
       
    Worksheets.Add after:=Worksheets(Worksheets.Count)                                  ' 末尾にシート追加
    ActiveSheet.Name = strFDName
       
        '----階層数制限 ----
        Limitation = Application.InputBox("管理階層数を整数で設定してください。（1,2,3・・・）" & vbCr & "（0=指定無）", Title:="管理階層数設定", Default:=0, Type:=1)
'        If Limitation = "" Then GoTo Exit_Handler
        If VarType(Limitation) = vbBoolean Then GoTo Exit_Handler
        
'Boolean型（False）が帰ってきたら，ラベル「Exit_Handler」へ
    
        Cells.Delete                                                                    '全セルの削除（白紙にする）
        Cells(1, 1) = strFD                                                             '親フォルダ名の書き込み
        
        '---- 階層の深さ，書き込み ----
        If Limitation < 1 Then
            Range("a2") = 0
            Limitation = 100
        Else
            Range("a2") = CInt(Limitation)                                              '引数をlong型（整数型）に変換
        End If
    End If
    
    sglT = Timer                                                                        '処理の経過時間を計測開始
    
    '---- セルにフォルダのパスを書き出す ----
    If Right(Cells(1, 1), 1) <> "\" Then
        strConnectionCharacter = "\"                                                        'ドライブ指定の時
    Else
        strConnectionCharacter = ""                                                         'フォルダ指定の時
    End If
    
    '---- サブフォルダの書き出し ----
    CountFolder = 1
    Call ExportSubfolders(strFD, strConnectionCharacter)                                    'サブフォルダの書き出し（再帰呼び出し）＆セルに分割
    
    '---- 同じ文字を省略する ----
    Call DeleteIfSame
    
    Range("a2") = ""
    Rows(2).Insert
    Range(Columns(2), Columns(FindLastColumn_MAX)).ColumnWidth = 20                         '複数列の幅を設定する
    
    Range("B2").Select
    
    For lngRetsu = 2 To FindLastColumn_MAX
        Cells(2, lngRetsu) = (lngRetsu - 1) & "階層"                                        '2行目・X列目のセルに階層を入力’
        Cells(2, lngRetsu).Interior.ColorIndex = 33 + lngRetsu                              ' 水色
    Next lngRetsu
    
    Application.StatusBar = "検索フォルダ数：　" & CountFolder - 1
    MsgBox "完了 " & Timer - sglT & " sec", vbInformation

Exit_Handler:
    Application.StatusBar = False                                                           'ここで元に戻す(Public変数)
    InterruptionFlg = False

End Sub

'-------------------------------------------------------------------------------
'サブフォルダの書き出し（再帰呼び出し）
'-------------------------------------------------------------------------------
Sub ExportSubfolders(strPath As String, strConnectionCharacter As String)

    Dim objFD As Object
    Dim strFD_NAME As String
    Dim varTmp As Variant
    Dim lngN As Long
    
    On Error Resume Next                                                        'エラーを無視して次の処理をする（システム関係のフォルダでうまく処理できない時に）
    Class = Class + 1
    
    With CreateObject("Scripting.FileSystemObject")                             'ファイルシステムオブジェクトを使用
        For Each objFD In .GetFolder(strPath).SubFolders
            
            '---- 中断処理 ----
            DoEvents
            If InterruptionFlg = False Then
                MsgBox "中断しました。", vbCritical
                Application.StatusBar = False
                End
            End If
            
            '----- 書き込み処理など ----
            strFD_NAME = objFD.Path
            strFD_NAME = Replace(strFD_NAME, Cells(1, 1) & strConnectionCharacter, "", , , vbTextCompare)
            
'↑親フォルダの名前を取り除いた「フォルダ名」にする vbTextCompare = 大文字小文字を区別しない
            
            varTmp = Split(strFD_NAME, "\")                                       'フォルダ名を\で分割して１次配列に格納（０から始まる）
            lngN = UBound(varTmp) '配列の上限（データ数－１）

            CountFolder = CountFolder + 1
            Cells(CountFolder, 2).Resize(1, lngN + 1).NumberFormatLocal = "@"     'セルの書式設定を文字列設定に
            Cells(CountFolder, 2).Resize(1, lngN + 1) = varTmp                            'セルの範囲を広げて一括で書き込む（セルの始まりとデータの始まりを一致させて処理)
'            Cells(CountFolder, 2) = strFD_NAME'書き込みテスト
            Application.StatusBar = "検索フォルダ数：　" & CountFolder - 1
            
            '---- 再帰呼び出し ----
            If Class < Limitation Then                                             '階層位置が制限より小さいなら
                Call ExportSubfolders(objFD.Path, strConnectionCharacter)                        'そのパスのサブフォルダを探す
            End If
            
        Next objFD
    End With
    
    Class = Class - 1
    
End Sub

'-------------------------------------------------------------------------------
'上の行と文字列が同じなら消す
'-------------------------------------------------------------------------------
Sub DeleteIfSame()

    Dim lngR As Long
    Dim lngC As Long
    Dim varTmp As Variant
    
    '---- 空行空列の削除（システムフォルダ関係の時に発生する） ----
    Call DeletingEmptyRowsColumns
    
    '---- 文字の消去 ----
    varTmp = Cells(1, 1).Resize(FindLastLine_MAX, FindLastColumn_MAX)               '最終行・最終列の関数は「Module6」にある
    
    If VarType(varTmp) = vbVariant + vbArray Then                                   '変数varTmpが，バリアント型で配列なら（１２＋８１９２＝８２０４←数値で比較する場合）
        For lngR = UBound(varTmp, 1) To LBound(varTmp, 1) + 1 Step -1               '最終行から2行目に向かって処理をする
            For lngC = LBound(varTmp, 2) To UBound(varTmp, 2)
                If varTmp(lngR, lngC) = varTmp(lngR - 1, lngC) Then varTmp(lngR, lngC) = Empty
                
'　""だと空文字を入力してしまうため
'Emptyはバリアント型だけの特別な空データで，数値だと０，文字列だと長さ０の文字列になる

            Next lngC
        Next lngR

        Cells(1, 1).Resize(FindLastLine_MAX, FindLastColumn_MAX) = varTmp
    End If
    
End Sub

'-------------------------------------------------------------------------------
'空行空列の削除
'-------------------------------------------------------------------------------
Sub DeletingEmptyRowsColumns()

    Dim lngR As Long
    Dim lngC As Long

    '---- 空行の削除 ----
    For lngR = FindLastLine_MAX To 2 Step -1
        If Application.WorksheetFunction.CountBlank(Rows(lngR)) = Columns.Count Then
'            MsgBox "lngR=" & lngR
            Rows(lngR).Delete
        End If
    Next lngR

    '---- 空列の削除 ----
    For lngC = FindLastColumn_MAX To 2 Step -1
        If Application.WorksheetFunction.CountBlank(Columns(lngC)) = Rows.Count Then
            Columns(lngC).Delete
'            MsgBox "lngC=" & lngC
        End If
    Next lngC

End Sub

'-------------------------------------------------------------------------------
'データの入っている最大最終行を求める
'-------------------------------------------------------------------------------
Function FindLastLine_MAX() As Long

    FindLastLine_MAX = ActiveSheet.UsedRange.Item(ActiveSheet.UsedRange.Count).Row          '使用している範囲での最後のセルの行を求める
    
End Function

'-------------------------------------------------------------------------------
'データの入っている最大最終列を求める
'-------------------------------------------------------------------------------
Function FindLastColumn_MAX() As Long

    FindLastColumn_MAX = ActiveSheet.UsedRange.Item(ActiveSheet.UsedRange.Count).Column     '使用している範囲の最後のセルの列を求める
    
End Function


