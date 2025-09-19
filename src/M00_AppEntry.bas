Attribute VB_Name = "M00_AppEntry"
'============================================================
' M00_AppEntry : �G���g���[�|�C���g/�t���[����
'============================================================
Option Explicit

Public farm_list As Variant
Public farm_name As String
Public start_date As Date
Public end_date As Date

Public Sub main()
    Dim start_time As Date, end_time As Date

    '�����s�m�F
    If Not ConfirmExecution() Then Exit Sub

    start_time = Time

    '���O�����F�V�[�g�������E����^�������t�擾
    Call PrepareSheetsAndDates

    '���f�[�^�擾�i����^�����^�ٓ����j
    Dim start_cows As Variant, end_cows As Variant, moved_cows As Variant
    Call LoadInitialCowLists(start_cows, end_cows, moved_cows)

    '����ID�����E�d������
    Dim cow_list As Variant
    cow_list = BuildUniqueCowList(start_cows, end_cows, moved_cows)
    If IsEmpty(cow_list) Or end_date = DateSerial(1999, 12, 31) Then
        MsgBox "���̓f�[�^������܂���B�I�����܂��B"
        Exit Sub
    End If

    '���q�ꖼ�̎擾�i�ŏ���1�����g���Ď��ʁj
    farm_name = GetFarmName(cow_list)

    '������E�����ɋL�^�������������X�N���C�s���O
    Call SearchInitialCows(cow_list)

    '�������o�����i����j�̋�������
    Call SearchGeneratedCows
    
    '���o�̓V�[�g�̍쐬�i�ڍׁE�W�v�j
    Call CreateOutputSheets

    '���W�v�V�[�g���`�E�v�Z������
    Call FormatOutputSheets

    '���I���ʒm�Ə������ԕ\��
    end_time = Time
    Call NotifyCompletion(start_time, end_time)
End Sub

Public Function ConfirmExecution() As Boolean
    ConfirmExecution = (MsgBox( _
        "�}�N�������s���܂����H" & vbCrLf & _
        "�i�z�菈������: 20�`30���j", _
        vbYesNo + vbQuestion, _
        "���s�m�F") = vbYes)

End Function

Public Sub NotifyCompletion(ByVal t0 As Date, ByVal t1 As Date)
    ActiveWorkbook.Save
    MsgBox "�������������܂���" & vbCrLf & "��������: " & Format(t1 - t0, "h:mm:ss")
End Sub

