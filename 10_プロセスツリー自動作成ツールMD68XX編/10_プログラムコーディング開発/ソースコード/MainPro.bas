Attribute VB_Name = "MainPro"
'**********************************************************************************************
' �}�ʔԍ��@                �FFSP01100800003
' ���́@                    �F�v���Z�X�c���[�����쐬�c�[��MD68XX�Ґ}��
' �\�t�g�E�F�A�Ǘ��ԍ��h�c�@�FFSS0101037
' �\�t�g�E�F�A���@          �F�v���Z�X�c���[�����쐬�c�[��MD68XX��
' ���W���[�����@            �FMainPro
' �@�\�T�v�@�@�@            �F�v���W�F�N�g�Ǘ����A�v���Z�X�c���[�̍쐬�Ɖ^�p�Ǘ�������������B
'
' ���������@�@�@            �F2019/07/24 �V�K T.Takayasu
'
' Copyright(C) 2019 Sanken Electric Co., Ltd. All Rights Reserved.
'**********************************************************************************************

Option Explicit

'====================================================================================================
' Main CreateNewFolder
'   \\Hka0-fspc\c\Users\GITCLONE\Documents\GIT_CLONE
'   ��ƃt�H���_ ---> ��
'====================================================================================================
' CONTENTS (�v���Z�X�c���[�����쐬�c�[���@�\���S�ە�)
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
'   (1)�ݒu�ꏊ�F�C��
'   (2)A1�Z���Ƀv���Z�X�c���[�̍쐬�ꏊ���w��i\\Hka0-fspc\c\Users\GITCLONE\Documents\GIT_CLONE�j
'   (3)A2�Z���Ƀv���Z�X�c���[�̍ő�K�w�����L��
'   (4)�y�@�\���S�v���Z�X�c���[�쐬�z�{�^����I���E���s
'   (5)�@�\���S�ۃv���Z�X�c���[�̕ύX���͕K���u�v���Z�X�c���[�����쐬�c�[���@�\���S�ەҁv���g�p���邱��
'======================================================================================================
'
'-------------------------------------------------------------------------------
' Dim Table
'-------------------------------------------------------------------------------

Public FolderCounter As Long      ' �t�H���_���̃J�E���^
Public InterruptionFlg As Boolean   ' ���f�����̃t���O


'-------------------------------------------------------------------------------
' CreateNewFolder
'-------------------------------------------------------------------------------

Sub CreateNewFolder()
' ---- ���V�K�t�H���_���쐬����v���O�������i�V�[�g�͂����炸�C��������Ŏ����j ----

    Dim lngR As Long
    Dim lngC As Long
    Dim vntTmp As Variant
    Dim strBuf As String
    Dim sngT As Single
    Dim lngMsg As Long

    On Error Resume Next                                                            ' Err.Number�Ƒg�ݍ��킹�Ďg��

    ' ---- DeleteEmptyRowsColumns ----
    Call DeleteEmptyRowsColumns
    
    InterruptionFlg = Not InterruptionFlg                                           ' ������True�ƂȂ�
    If InterruptionFlg = True Then
        ' ---- �e�t�H���_�̐ݒ� ----
        If Cells(1, 1) <> "" Then
            lngMsg = MsgBox("A1�ɋL������Ă���t�H���_�Ƀc���[���쐬���܂��B", vbQuestion + vbYesNoCancel) ' ��
            If lngMsg = vbCancel Then GoTo Exit_Handler                                     ' ��lngMsg��Cancel�̎�
        End If
        
        If Cells(1, 1) = "" Or lngMsg = vbNo Then                                   ' A1���󂩁C��lngMsg��No�̂Ƃ�
            MsgBox "�ۑ��t�H���_���w�肳��Ă��܂���B"
            GoTo Exit_Handler
        End If
        
        ' ---- �֎~�����̃`�F�b�N ----
        If CheckForProhibitedCharacters.Address <> "$A$1" Then                      ' �Ȃ���΁C$A$1���Ԃ�
            CheckForProhibitedCharacters.Select
            MsgBox "���̃Z���Ɏg�p�֎~�̕������܂܂�Ă��܂��B" & vbLf & CheckForProhibitedCharacters.Address, vbCritical
            GoTo Exit_Handler
        End If
        
    End If

    ' ---- FillInTheBlanks ----
    vntTmp = Cells(1, 1).Resize(LastRow_MAX, LastColumn_MAX)                        ' �z��Ƃ��Ċi�[
    vntTmp = FillInTheBlanks(vntTmp)                                                ' �ϐ����
    
    ' ---- ��CreateParentFolder�� ----
    FolderCounter = 0
    Call CreateParentFolder(Cells(1, 1))
    
    sngT = Timer
    ' ---- ���V�K�t�H���_�̍쐬�� ----
    For lngR = 2 To LastRow_MAX
        strBuf = vntTmp(1, 1)

        ' ---- �쐬 ----
        For lngC = 2 To LastColumn(lngR)
            strBuf = strBuf & "\" & vntTmp(lngR, lngC)                             ' ��������Ȃ���
            
            Err.Number = 0
            MkDir (strBuf) ' �t�H���_�̍쐬
            If Err.Number = 0 Then FolderCounter = FolderCounter + 1                ' �쐬�ł�����i�G���[����Ȃ���΁j�J�E���^���{�P
            Application.StatusBar = "�쐬�t�H���_���F�@" & FolderCounter

            DoEvents

            If InterruptionFlg = False Then
                MsgBox "���f���܂����B", vbInformation
                GoTo Exit_Handler
            End If
        Next lngC
    Next lngR

    If FolderCounter > 0 Then
        MsgBox "�����@" & Timer - sngT & " sec", vbInformation
    Else
        MsgBox "�t�H���_���쐬�ł��܂���ł����B" & vbLf & "���̂悤�Ȍ������l�����܂��B" & vbLf & vbLf & "�E�h���C�u���Ȃ��B" & vbLf & "�E�����t�H���_�����ɂ���B�@�Ȃ�", vbCritical
    End If
    
    ActiveSheet.Hyperlinks.Delete                                                   ' �n�C�p�[�����N����������

Exit_Handler:
    Application.StatusBar = False
    InterruptionFlg = False                                                         ' �����Ō��ɖ߂�(Public�ϐ�)

End Sub

'-------------------------------------------------------------------------------
' DeleteEmptyMatrix
'-------------------------------------------------------------------------------

Sub DeleteEmptyMatrix()

    Call DeleteEmptyRowsColumns
    
    MsgBox "����", vbInformation

End Sub

'-------------------------------------------------------------------------------
' DeleteEmptyRowsColumns
'-------------------------------------------------------------------------------

Sub DeleteEmptyRowsColumns()

    Dim lngR As Long
    Dim lngC As Long

    ' ---- ��s�̍폜 ----
    For lngR = LastRow_MAX To 2 Step -1
        If Application.WorksheetFunction.CountBlank(Rows(lngR)) = Columns.Count Then
            Rows(lngR).Delete
        End If
    Next lngR

    ' ---- ���̍폜 ----
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
    
    ' ---- �t�@�C���V�X�e���I�u�W�F�N�g���g���āC�e�t�H���_�̗L���𒲂ׂ� ----Dir(strT, vbDirectory)�ł́C�h���C�u���󂾂ƌ��o�ł��Ȃ��̂�
    If objFSO.FolderExists(strFolder) = False Then
        lngMsg = MsgBox("�e�t�H���_�����݂��܂���B" & vbCr & "�V�K�ō���ď����𑱂��܂����H", vbYesNo + vbQuestion)
    End If

    ' ---- CreateParentFolder���������� ----
    If lngMsg = vbYes Then                                                  ' ���V�K�쐬��
    
        ' ---- MkDir�ō�邽�߂̉����� ----
        If Left(strBuf, 2) = "\\" Then
            strBuf = Right(strBuf, Len(strBuf) - 2)                         ' �l�b�g���[�N�Ή��i�擪��\\����菜���j
            strNetwork = "\\"
        End If
        
        vntTmp = Split(strBuf, "\")
        strT = strNetwork & vntTmp(0)                                       ' �l�b�g���[�N�Ή��i�擪��\\��t����j
        
        ' ---- �쐬���� ----
        For lngN = LBound(vntTmp) + 1 To UBound(vntTmp)
            strT = strT & "\" & vntTmp(lngN)
    
            Err.Number = 0
            MkDir (strT)                                                    ' �t�H���_�̍쐬
            If Err.Number = 0 Then FolderCounter = FolderCounter + 1        ' �쐬�ł�����i�G���[����Ȃ���΁j�J�E���^���{�P
    
            Application.StatusBar = "�쐬�t�H���_���F�@" & FolderCounter
        Next lngN

        
    ElseIf lngMsg = vbNo Then                                               ' �����f��
        InterruptionFlg = False                                             ' �����Ō��ɖ߂�(Public�ϐ�)
        Application.StatusBar = False

        End
    End If
    
End Sub

'-------------------------------------------------------------------------------
' CheckForProhibitedCharacters
'-------------------------------------------------------------------------------

Function CheckForProhibitedCharacters() As Range

' ---- �֎~�����̃`�F�b�N ----

    Dim lngN As Long
    Dim vntBan As Variant
    Dim objRg As Range                                                      ' ���ԒZ�k�̂���
    
    vntBan = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    
    For Each objRg In Cells(1, 2).Resize(LastRow_MAX, LastColumn_MAX - 1)
        For lngN = LBound(vntBan) To UBound(vntBan)
            If InStr(objRg.Value, vntBan(lngN)) > 0 Then                    ' �֎~����������Έʒu��Ԃ��̂�
                Set CheckForProhibitedCharacters = objRg
                Exit Function
            End If
        Next lngN
    Next objRg

    Set CheckForProhibitedCharacters = Cells(1, 1)                          ' �֎~�������Ȃ���΁C$A$1��Ԃ�

End Function

'-------------------------------------------------------------------------------
' FillInTheBlanks
'-------------------------------------------------------------------------------

Function FillInTheBlanks(vntTmp As Variant) As Variant

' ---- �Z���Ō����Ƃ���́CFillInTheBlanks ----

    Dim lngR As Long
    Dim lngC As Long
    
    ' ---- 2��ڂ܂ł̏��� ----
    For lngC = 1 To 2
        For lngR = 2 To LastRow_MAX
            If vntTmp(lngR, lngC) = "" Then
                vntTmp(lngR, lngC) = vntTmp(lngR - 1, lngC)
            End If
        Next lngR
    Next lngC
    
    ' ---- 3��ڈȍ~�̏���----
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
' ---- ���̗��LastRow�����߂� ----

    Dim lngR As Long
    
    lngR = Cells(Rows.Count, last_column).End(xlUp).Row                         ' �s������T��
    If Cells(1, last_column) = "" And lngR = 1 Then lngR = 0

    LastRow = lngR

End Function

'-------------------------------------------------------------------------------
' LastColumn
'-------------------------------------------------------------------------------

Function LastColumn(last_row As Long) As Long
' ---- ���̍s��LastColumn�����߂� ----

    Dim lngC As Long
    
    lngC = Cells(last_row, Columns.Count).End(xlToLeft).Column                  ' �񖖂���T��
    If Cells(last_row, 1) = "" And lngC = 1 Then lngC = 0

    LastColumn = lngC

End Function

'-------------------------------------------------------------------------------
' LastRow_MAX
'-------------------------------------------------------------------------------

Function LastRow_MAX() As Long
' ---- �f�[�^�̓����Ă���ő�LastRow�����߂� ----

    LastRow_MAX = ActiveSheet.UsedRange.Item(ActiveSheet.UsedRange.Count).Row   ' �g�p���Ă���͈͂ł̍Ō�̃Z���̍s�����߂�
    
End Function

'-------------------------------------------------------------------------------
' LastColumn_MAX
'-------------------------------------------------------------------------------

Function LastColumn_MAX() As Long
' ---- �f�[�^�̓����Ă���ő�LastColumn�����߂� ----

    LastColumn_MAX = Range("A2").Value                                          ' �ő�K�w��A2�̒l�ɌŒ�
    LastColumn_MAX = LastColumn_MAX + 1

End Function
