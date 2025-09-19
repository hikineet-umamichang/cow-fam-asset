Attribute VB_Name = "M00_AppEntry"
Option Explicit

Public farm_list As Variant
Public farm_name As String
Public start_date As Date
Public end_date As Date

Sub MAIN()
    Dim start_time As Date, end_time As Date

    '�����s�m�F
    If Not ConfirmExecution Then Exit Sub

    start_time = Time

    '���O�����F�V�[�g�������E����^�������t�擾
    Call PrepareSheetsAndDates

    '���f�[�^�擾�i����^�����^�ٓ����j
    Dim start_cows As Variant, end_cows As Variant, moved_cows As Variant
    Call LoadInitialCowLists(start_cows, end_cows, moved_cows)

    '���o�̓V�[�g�̍쐬�i�ڍׁE�W�v�j
    Call CreateOutputSheets

    '����ID�����E�d������
    Dim cow_list As Variant
    cow_list = BuildUniqueCowList(start_cows, end_cows, moved_cows)
    If IsEmpty(cow_list) Or end_date = "1999/12/31" Then
        MsgBox "���̓f�[�^������܂���B�I�����܂��B"
        Exit Sub
    End If

    '���q�ꖼ�̎擾�i�ŏ���1�����g���Ď��ʁj
    farm_name = GetFarmName(cow_list)

    '������E�����ɋL�^�������������X�N���C�s���O
    Call SearchInitialCows(cow_list)

    '�������o�����i����j�̋�������
    Call SearchGeneratedCows

    '���W�v�V�[�g���`�E�v�Z������
    Call FormatOutputSheets

    '���I���ʒm�Ə������ԕ\��
    end_time = Time
    Call NotifyCompletion(start_time, end_time)
End Sub


'------------------------------------------------------
' ���[�U�[�Ƀ}�N�����s�ۂ��m�F����
' ���s����ꍇ�� True�A�L�����Z���̏ꍇ�� False ��Ԃ�
'------------------------------------------------------
Function ConfirmExecution() As Boolean
    Dim userResponse As VbMsgBoxResult
    
    userResponse = MsgBox( _
        "�}�N�������s���܂����H" & vbCrLf & _
        "�i�z�菈������: 20?30���j", _
        vbYesNo + vbQuestion, _
        "���s�m�F")
    
    If userResponse = vbYes Then
        ConfirmExecution = True
    Else
        ConfirmExecution = False
    End If
End Function




