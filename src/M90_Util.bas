Attribute VB_Name = "M90_Util"
'============================================================
' M90_Util : ���ʃ��[�e�B���e�B
'============================================================
Option Explicit

' ���O�i�ȈՁj
Public Sub LogInfo(ByVal msg As String): Debug.Print "[INFO] " & msg: End Sub
Public Sub LogWarn(ByVal msg As String): Debug.Print "[WARN] " & msg: End Sub
Public Sub LogError(ByVal msg As String): Debug.Print "[ERROR] " & msg: End Sub

'------------------------------------------------------------
' ��z�񔻒�iArray() ���܂߈��S�ɔ���j
'------------------------------------------------------------
Public Function IsEmptyArray(ByVal v As Variant) As Boolean
    Dim lb As Long, ub As Long
    If Not IsArray(v) Then
        IsEmptyArray = True
        Exit Function
    End If
    On Error Resume Next
    lb = LBound(v)
    ub = UBound(v)
    If Err.number <> 0 Then
        IsEmptyArray = True
    Else
        IsEmptyArray = (ub < lb)
    End If
    On Error GoTo 0
End Function

'------------------------------------------------------------
' Range/Variant �� 1D�x�N�^�ɐ��K���i��E�󔒂����O�j
'  - 2D(�c1��/��1�s)�ɂ��Ή�
'  - �P��l�̏ꍇ�͒P��v�f�x�N�^��
'------------------------------------------------------------
Public Function ToVector(ByVal v As Variant) As Variant
    Dim tmp() As Variant
    Dim r As Long, c As Long, i As Long
    Dim n As Long

    If IsEmpty(v) Then
        ToVector = Array(): Exit Function
    End If

    If IsArray(v) Then
        Dim lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long
        On Error Resume Next
        lb1 = LBound(v, 1): ub1 = UBound(v, 1)
        lb2 = LBound(v, 2): ub2 = UBound(v, 2)
        If Err.number <> 0 Then
            ' 1�����z��Ƃ��Ĉ���
            Err.Clear
            lb1 = LBound(v): ub1 = UBound(v)
            ReDim tmp(0 To 0)
            For i = lb1 To ub1
                If Trim(CStr(v(i))) <> "" Then
                    If n = 0 Then ReDim tmp(0 To 0) Else ReDim Preserve tmp(0 To n)
                    tmp(n) = Trim(CStr(v(i))): n = n + 1
                End If
            Next
            If n = 0 Then ToVector = Array() Else ToVector = tmp
            Exit Function
        End If
        On Error GoTo 0

        ' 2�����iRange.Value�j��z��
        ReDim tmp(0 To 0)
        For r = lb1 To ub1
            For c = lb2 To ub2
                If Trim(CStr(v(r, c))) <> "" Then
                    If n = 0 Then ReDim tmp(0 To 0) Else ReDim Preserve tmp(0 To n)
                    tmp(n) = Trim(CStr(v(r, c))): n = n + 1
                End If
            Next c
        Next r
        If n = 0 Then ToVector = Array() Else ToVector = tmp
    Else
        ' �P��l
        If Trim(CStr(v)) = "" Then
            ToVector = Array()
        Else
            ToVector = Array(Trim(CStr(v)))
        End If
    End If
End Function

'------------------------------------------------------------
' 1D�x�N�^���un�s�~1��v��2D�z��ɕϊ��iRange�݊��j
'------------------------------------------------------------
Public Function MakeColumn2D(ByVal vec As Variant) As Variant
    If IsEmptyArray(vec) Then
        MakeColumn2D = Empty
        Exit Function
    End If
    Dim n As Long, i As Long
    n = UBound(vec) - LBound(vec) + 1
    Dim Arr() As Variant
    ReDim Arr(1 To n, 1 To 1)
    For i = 1 To n
        Arr(i, 1) = vec(LBound(vec) + i - 1)
    Next
    MakeColumn2D = Arr
End Function

'------------------------------------------------------------
' �w��Z�����牺�Ɍ������Ēl��ǂݎ��A���o���݂̂Ȃ��z���Ԃ�
' ��FGetColumnValues(ws, "A2")
'------------------------------------------------------------
Public Function GetColumnValues(ByVal ws As Worksheet, ByVal startCellAddress As String) As Variant
    Dim startCell As Range, lastCell As Range, dataRange As Range
    Set startCell = ws.Range(startCellAddress)
    Set lastCell = ws.Cells(ws.Rows.Count, startCell.Column).End(xlUp)

    If lastCell.Row < startCell.Row Then
        GetColumnValues = Array()
        Exit Function
    End If

    Set dataRange = ws.Range(startCell, lastCell)

    ' 1�Z�������i= ���o���̉\���j����� �� ���ꂾ���Ŕz��
    If dataRange.Rows.Count = 1 Then
        If Trim(CStr(dataRange.Value2)) = "" Then
            GetColumnValues = Array()
        Else
            GetColumnValues = Array(Trim(CStr(dataRange.Value2)))
        End If
    Else
        GetColumnValues = dataRange.value
    End If
End Function

Public Function ArrayColumn(ByVal rng As Range) As Variant
    ArrayColumn = rng.value
End Function

Public Sub ExportAllModules(ByVal exportPath As String)
    Dim vbComp As VBIDE.VBComponent
    Dim fso As Object
    Dim fileName As String

    If Len(Dir(exportPath, vbDirectory)) = 0 Then MkDir exportPath

    For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_StdModule: fileName = exportPath & "\" & vbComp.Name & ".bas"
            Case vbext_ct_ClassModule: fileName = exportPath & "\" & vbComp.Name & ".cls"
            Case vbext_ct_MSForm: fileName = exportPath & "\" & vbComp.Name & ".frm"
            Case Else: fileName = ""
        End Select

        If fileName <> "" Then
            On Error Resume Next
            Kill fileName        ' �������폜���ď㏑��
            On Error GoTo 0
            vbComp.Export fileName
        End If
    Next vbComp
End Sub


