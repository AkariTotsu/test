Attribute VB_Name = "MainPro"
'**********************************************************************************************
' �}�ʔԍ��@                �FFSP01110800001
' ���́@                    �F�v���Z�X�c���[�\�������鉻
' �\�t�g�E�F�A�Ǘ��ԍ��h�c�@�FFSS0101038
' �\�t�g�E�F�A���@          �F�v���Z�X�c���[�\�������鉻
' ���W���[�����@            �FMainPro
' �@�\�T�v�@�@�@            �F�v���Z�X�c���[�̍\�����G�N�Z���V�[�g��Ɏ����쐬�������鉻����B
'
' ���������@�@�@            �F2019/08/16 �V�K T.Takayasu
'
' Copyright(C) 2019 Sanken Electric Co., Ltd. All Rights Reserved.
'**********************************************************************************************

Option Explicit

'====================================================================================================
' Main CreateNewFolder
'   \\Hka0-fspc\c\Users\GITCLONE\Documents\GIT_CLONE
'   ��ƃt�H���_ ---> ��
'====================================================================================================
' CONTENTS (�v���Z�X�c���[�\�������鉻)
'   ExportFolderStructure
'   ExportSubfolders
'   DeleteIfSame
'   DeletingEmptyRowsColumns
'   FindLastLine_MAX
'   FindLastColumn_MAX
'=====================================================================================================
' DEPLOYMENT
'   (1)�ݒu�ꏊ�F�C��
'   (2)B3�Z���Ƀv���Z�X�c���[�̏����t�H���_���w��iC:\Users\takayasu-toshiyuki\Documents\�v���Z�X�c���[�쐬�e�X�g�p\�j
'   (3)�y�v���Z�X�c���[�\�������鉻�z�{�^����I���E���s
'======================================================================================================
'
'-------------------------------------------------------------------------------
' Dim Table
'-------------------------------------------------------------------------------

Public CountFolder As Long                                  '�t�H���_���̃J�E���^
Public InterruptionFlg As Boolean                           '���f�����̃t���O
Dim Class As Long                                           '�K�w�̃J�E���^
Dim Limitation As Variant                                   '�K�w�̐[������

'-------------------------------------------------------------------------------
' �t�H���_�̍\���������o���v���O����
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

    InterruptionFlg = Not InterruptionFlg                                               '�����l�́C������True�ɂȂ�
    
    Sheets("INDEX").Select

    dataFolder = Range("B3").Value                                                      '�����t�H���_

    If InterruptionFlg = True Then
        '---- ��������t�H���_��I������ ----
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "�v���Z�X�c���[�\����ǂݍ��ރt�H���_��I�����ĉ������B�i���s���ɁC������x�s���ƒ��f���܂��B�j"
            .InitialFileName = dataFolder                                               '�����t�H���_���w�肷��
            If .Show = False Then GoTo Exit_Handler                                     '�_�C�A���O�{�b�N�X��\��
            strFD = .SelectedItems(1)                                                   '�t�H���_�����i�[
        End With
        
    strFDName = Right(strFD, InStr(StrReverse(strFD), "\") - 1)                         '�ŉ��w�̃t�H���_�����擾
       
    Worksheets.Add after:=Worksheets(Worksheets.Count)                                  ' �����ɃV�[�g�ǉ�
    ActiveSheet.Name = strFDName
       
        '----�K�w������ ----
        Limitation = Application.InputBox("�Ǘ��K�w���𐮐��Őݒ肵�Ă��������B�i1,2,3�E�E�E�j" & vbCr & "�i0=�w�薳�j", Title:="�Ǘ��K�w���ݒ�", Default:=0, Type:=1)
'        If Limitation = "" Then GoTo Exit_Handler
        If VarType(Limitation) = vbBoolean Then GoTo Exit_Handler
        
'Boolean�^�iFalse�j���A���Ă�����C���x���uExit_Handler�v��
    
        Cells.Delete                                                                    '�S�Z���̍폜�i�����ɂ���j
        Cells(1, 1) = strFD                                                             '�e�t�H���_���̏�������
        
        '---- �K�w�̐[���C�������� ----
        If Limitation < 1 Then
            Range("a2") = 0
            Limitation = 100
        Else
            Range("a2") = CInt(Limitation)                                              '������long�^�i�����^�j�ɕϊ�
        End If
    End If
    
    sglT = Timer                                                                        '�����̌o�ߎ��Ԃ��v���J�n
    
    '---- �Z���Ƀt�H���_�̃p�X�������o�� ----
    If Right(Cells(1, 1), 1) <> "\" Then
        strConnectionCharacter = "\"                                                        '�h���C�u�w��̎�
    Else
        strConnectionCharacter = ""                                                         '�t�H���_�w��̎�
    End If
    
    '---- �T�u�t�H���_�̏����o�� ----
    CountFolder = 1
    Call ExportSubfolders(strFD, strConnectionCharacter)                                    '�T�u�t�H���_�̏����o���i�ċA�Ăяo���j���Z���ɕ���
    
    '---- �����������ȗ����� ----
    Call DeleteIfSame
    
    Range("a2") = ""
    Rows(2).Insert
    Range(Columns(2), Columns(FindLastColumn_MAX)).ColumnWidth = 20                         '������̕���ݒ肷��
    
    Range("B2").Select
    
    For lngRetsu = 2 To FindLastColumn_MAX
        Cells(2, lngRetsu) = (lngRetsu - 1) & "�K�w"                                        '2�s�ځEX��ڂ̃Z���ɊK�w����́f
        Cells(2, lngRetsu).Interior.ColorIndex = 33 + lngRetsu                              ' ���F
    Next lngRetsu
    
    Application.StatusBar = "�����t�H���_���F�@" & CountFolder - 1
    MsgBox "���� " & Timer - sglT & " sec", vbInformation

Exit_Handler:
    Application.StatusBar = False                                                           '�����Ō��ɖ߂�(Public�ϐ�)
    InterruptionFlg = False

End Sub

'-------------------------------------------------------------------------------
'�T�u�t�H���_�̏����o���i�ċA�Ăяo���j
'-------------------------------------------------------------------------------
Sub ExportSubfolders(strPath As String, strConnectionCharacter As String)

    Dim objFD As Object
    Dim strFD_NAME As String
    Dim varTmp As Variant
    Dim lngN As Long
    
    On Error Resume Next                                                        '�G���[�𖳎����Ď��̏���������i�V�X�e���֌W�̃t�H���_�ł��܂������ł��Ȃ����Ɂj
    Class = Class + 1
    
    With CreateObject("Scripting.FileSystemObject")                             '�t�@�C���V�X�e���I�u�W�F�N�g���g�p
        For Each objFD In .GetFolder(strPath).SubFolders
            
            '---- ���f���� ----
            DoEvents
            If InterruptionFlg = False Then
                MsgBox "���f���܂����B", vbCritical
                Application.StatusBar = False
                End
            End If
            
            '----- �������ݏ����Ȃ� ----
            strFD_NAME = objFD.Path
            strFD_NAME = Replace(strFD_NAME, Cells(1, 1) & strConnectionCharacter, "", , , vbTextCompare)
            
'���e�t�H���_�̖��O����菜�����u�t�H���_���v�ɂ��� vbTextCompare = �啶������������ʂ��Ȃ�
            
            varTmp = Split(strFD_NAME, "\")                                       '�t�H���_����\�ŕ������ĂP���z��Ɋi�[�i�O����n�܂�j
            lngN = UBound(varTmp) '�z��̏���i�f�[�^���|�P�j

            CountFolder = CountFolder + 1
            Cells(CountFolder, 2).Resize(1, lngN + 1).NumberFormatLocal = "@"     '�Z���̏����ݒ�𕶎���ݒ��
            Cells(CountFolder, 2).Resize(1, lngN + 1) = varTmp                            '�Z���͈̔͂��L���Ĉꊇ�ŏ������ށi�Z���̎n�܂�ƃf�[�^�̎n�܂����v�����ď���)
'            Cells(CountFolder, 2) = strFD_NAME'�������݃e�X�g
            Application.StatusBar = "�����t�H���_���F�@" & CountFolder - 1
            
            '---- �ċA�Ăяo�� ----
            If Class < Limitation Then                                             '�K�w�ʒu��������菬�����Ȃ�
                Call ExportSubfolders(objFD.Path, strConnectionCharacter)                        '���̃p�X�̃T�u�t�H���_��T��
            End If
            
        Next objFD
    End With
    
    Class = Class - 1
    
End Sub

'-------------------------------------------------------------------------------
'��̍s�ƕ����񂪓����Ȃ����
'-------------------------------------------------------------------------------
Sub DeleteIfSame()

    Dim lngR As Long
    Dim lngC As Long
    Dim varTmp As Variant
    
    '---- ��s���̍폜�i�V�X�e���t�H���_�֌W�̎��ɔ�������j ----
    Call DeletingEmptyRowsColumns
    
    '---- �����̏��� ----
    varTmp = Cells(1, 1).Resize(FindLastLine_MAX, FindLastColumn_MAX)               '�ŏI�s�E�ŏI��̊֐��́uModule6�v�ɂ���
    
    If VarType(varTmp) = vbVariant + vbArray Then                                   '�ϐ�varTmp���C�o���A���g�^�Ŕz��Ȃ�i�P�Q�{�W�P�X�Q���W�Q�O�S�����l�Ŕ�r����ꍇ�j
        For lngR = UBound(varTmp, 1) To LBound(varTmp, 1) + 1 Step -1               '�ŏI�s����2�s�ڂɌ������ď���������
            For lngC = LBound(varTmp, 2) To UBound(varTmp, 2)
                If varTmp(lngR, lngC) = varTmp(lngR - 1, lngC) Then varTmp(lngR, lngC) = Empty
                
'�@""���Ƌ󕶎�����͂��Ă��܂�����
'Empty�̓o���A���g�^�����̓��ʂȋ�f�[�^�ŁC���l���ƂO�C�����񂾂ƒ����O�̕�����ɂȂ�

            Next lngC
        Next lngR

        Cells(1, 1).Resize(FindLastLine_MAX, FindLastColumn_MAX) = varTmp
    End If
    
End Sub

'-------------------------------------------------------------------------------
'��s���̍폜
'-------------------------------------------------------------------------------
Sub DeletingEmptyRowsColumns()

    Dim lngR As Long
    Dim lngC As Long

    '---- ��s�̍폜 ----
    For lngR = FindLastLine_MAX To 2 Step -1
        If Application.WorksheetFunction.CountBlank(Rows(lngR)) = Columns.Count Then
'            MsgBox "lngR=" & lngR
            Rows(lngR).Delete
        End If
    Next lngR

    '---- ���̍폜 ----
    For lngC = FindLastColumn_MAX To 2 Step -1
        If Application.WorksheetFunction.CountBlank(Columns(lngC)) = Rows.Count Then
            Columns(lngC).Delete
'            MsgBox "lngC=" & lngC
        End If
    Next lngC

End Sub

'-------------------------------------------------------------------------------
'�f�[�^�̓����Ă���ő�ŏI�s�����߂�
'-------------------------------------------------------------------------------
Function FindLastLine_MAX() As Long

    FindLastLine_MAX = ActiveSheet.UsedRange.Item(ActiveSheet.UsedRange.Count).Row          '�g�p���Ă���͈͂ł̍Ō�̃Z���̍s�����߂�
    
End Function

'-------------------------------------------------------------------------------
'�f�[�^�̓����Ă���ő�ŏI������߂�
'-------------------------------------------------------------------------------
Function FindLastColumn_MAX() As Long

    FindLastColumn_MAX = ActiveSheet.UsedRange.Item(ActiveSheet.UsedRange.Count).Column     '�g�p���Ă���͈͂̍Ō�̃Z���̗�����߂�
    
End Function


