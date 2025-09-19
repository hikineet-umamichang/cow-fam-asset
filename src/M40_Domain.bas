Attribute VB_Name = "M40_Domain"
'============================================================
' M40_Domain
' �ړI : ��ID�����E�`�F�b�N�f�B�W�b�g�E�Ɩ����W�b�N
' �ˑ� : �Ȃ��i���ʂ�UTL�̂݁j
' �� : GenerateCowID, AddCheckDigit, ClassifyMovement ...
'============================================================
Option Explicit


' �w�肳�ꂽ�̎��ʔԍ���interval�������������A�V�����`�F�b�N�f�B�W�b�g�t���ԍ���Ԃ�
Function GenerateCowID(ByVal cowIDNumber As String, ByVal interval As Long) As String
    Dim farmCode As Long
    Dim serialNumber As Long
    Dim newSerialBase As Long
    Dim baseWithoutCheckDigit As Long

    ' �擪5�� = �q��R�[�h�A����4�� = �A��
    farmCode = CLng(Left(cowIDNumber, 5))
    serialNumber = CLng(Mid(cowIDNumber, 6, 4))

    ' �A�Ԃ�interval���������Z
    newSerialBase = serialNumber + interval
    
    ' 4���Ń��[���I�[�o�[�i0000 �� 9999�j����ꍇ�̕␳�i���K�v�Ȃ�j
    If newSerialBase > 9999 Then newSerialBase = newSerialBase - 10000
    If newSerialBase < 0 Then newSerialBase = newSerialBase + 10000

    ' �`�F�b�N�f�B�W�b�g��������9���̐������쐬
    baseWithoutCheckDigit = farmCode * 10000 + newSerialBase

    ' �`�F�b�N�f�B�W�b�g��t���ĕԂ�
    GenerateCowID = AddCheckDigit(baseWithoutCheckDigit)
End Function


' �`�F�b�N�f�B�W�b�g��t�����̎��ʔԍ��i10���j��Ԃ�
Function AddCheckDigit(ByVal number As Long) As String
    Dim evenSum As Long, oddSum As Long
    Dim numberStr As String
    Dim i As Integer

    numberStr = Format(number, "000000000") ' 9����0����

    ' ����Ƌ������ŕ����ĉ��Z�i1-based index�j
    For i = 1 To 9
        If i Mod 2 = 0 Then
            evenSum = evenSum + CInt(Mid(numberStr, i, 1))
        Else
            oddSum = oddSum + CInt(Mid(numberStr, i, 1))
        End If
    Next i

    ' �`�F�b�N�f�B�W�b�g�v�Z�i���W�����X10�␔�j
    Dim checkDigit As Integer
    checkDigit = (10 - ((oddSum * 3 + evenSum) Mod 10)) Mod 10

    AddCheckDigit = numberStr & CStr(checkDigit)
End Function

