Attribute VB_Name = "M20_Input"
'============================================================
' M20_Input : 入力読み取り/統合/重複排除
'============================================================
Option Explicit

Public Sub LoadInitialCowLists(ByRef start_cows As Variant, ByRef end_cows As Variant, ByRef moved_cows As Variant)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("入力シート")

    start_cows = GetColumnValues(ws, "A2")  ' M90_Util
    end_cows = GetColumnValues(ws, "C2")
    moved_cows = GetColumnValues(ws, "E2")
End Sub

'------------------------------------------------------------
' 3つのリストを統合し、重複を除いた「n行×1列」2D配列を返す
' すべて空なら Empty を返す
'------------------------------------------------------------
Public Function BuildUniqueCowList(ByVal start_cows As Variant, ByVal end_cows As Variant, ByVal moved_cows As Variant) As Variant
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare  ' 文字列キーの大文字小文字を無視

    ' 正規化して追加
    AddVectorToDict ToVector(start_cows), dict
    AddVectorToDict ToVector(end_cows), dict
    AddVectorToDict ToVector(moved_cows), dict

    If dict.Count = 0 Then
        BuildUniqueCowList = Empty
        Exit Function
    End If

    ' Dictionary → 1Dベクタ → n×1 の2D
    Dim keys As Variant
    keys = dict.keys
    BuildUniqueCowList = MakeColumn2D(keys)  ' M90_Util
End Function

' 内部：ベクタを辞書に投入（空白・重複は自然にスキップ）
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
