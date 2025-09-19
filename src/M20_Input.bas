Attribute VB_Name = "M20_Input"
'============================================================
' M20_Input : ���͓ǂݎ��/����/�d���r��
'============================================================
Option Explicit

Public Sub LoadInitialCowLists(ByRef start_cows As Variant, ByRef end_cows As Variant, ByRef moved_cows As Variant)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("���̓V�[�g")

    start_cows = GetColumnValues(ws, "A2")  ' M90_Util
    end_cows = GetColumnValues(ws, "C2")
    moved_cows = GetColumnValues(ws, "E2")
End Sub

'------------------------------------------------------------
' 3�̃��X�g�𓝍����A�d�����������un�s�~1��v2D�z���Ԃ�
' ���ׂċ�Ȃ� Empty ��Ԃ�
'------------------------------------------------------------
Public Function BuildUniqueCowList(ByVal start_cows As Variant, ByVal end_cows As Variant, ByVal moved_cows As Variant) As Variant
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare  ' ������L�[�̑啶���������𖳎�

    ' ���K�����Ēǉ�
    AddVectorToDict ToVector(start_cows), dict
    AddVectorToDict ToVector(end_cows), dict
    AddVectorToDict ToVector(moved_cows), dict

    If dict.Count = 0 Then
        BuildUniqueCowList = Empty
        Exit Function
    End If

    ' Dictionary �� 1D�x�N�^ �� n�~1 ��2D
    Dim keys As Variant
    keys = dict.keys
    BuildUniqueCowList = MakeColumn2D(keys)  ' M90_Util
End Function

' �����F�x�N�^�������ɓ����i�󔒁E�d���͎��R�ɃX�L�b�v�j
Private Sub AddVectorToDict(ByVal vec As Variant, ByRef dict As Object)
    If IsEmptyArray(vec) Then Exit Sub
    Dim i As Long, k As String
    For i = LBound(vec) To UBound(vec)
        k = Trim(CStr(vec(i)))
        If Len(k) > 0 Then
            If Not dict.Exists(k) Then dict.Add k, True
        End If
    Next i
End Sub
