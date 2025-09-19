Attribute VB_Name = "M90_Util"
'============================================================
' M90_Util : 共通ユーティリティ
'============================================================
Option Explicit

' ログ（簡易）
Public Sub LogInfo(ByVal msg As String): Debug.Print "[INFO] " & msg: End Sub
Public Sub LogWarn(ByVal msg As String): Debug.Print "[WARN] " & msg: End Sub
Public Sub LogError(ByVal msg As String): Debug.Print "[ERROR] " & msg: End Sub

'------------------------------------------------------------
' 空配列判定（Array() を含め安全に判定）
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
' Range/Variant を 1Dベクタに正規化（空・空白を除外）
'  - 2D(縦1列/横1行)にも対応
'  - 単一値の場合は単一要素ベクタ化
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
            ' 1次元配列として扱う
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

        ' 2次元（Range.Value）を想定
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
        ' 単一値
        If Trim(CStr(v)) = "" Then
            ToVector = Array()
        Else
            ToVector = Array(Trim(CStr(v)))
        End If
    End If
End Function

'------------------------------------------------------------
' 1Dベクタを「n行×1列」の2D配列に変換（Range互換）
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
' 指定セルから下に向かって値を読み取り、見出しのみなら空配列を返す
' 例：GetColumnValues(ws, "A2")
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

    ' 1セルだけ（= 見出しの可能性）かつ非空 → それだけで配列化
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
            Kill fileName        ' 既存を削除して上書き
            On Error GoTo 0
            vbComp.Export fileName
        End If
    Next vbComp
End Sub


