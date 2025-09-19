Attribute VB_Name = "M31_Parse"
'============================================================
' M31_Parse : HTML解析 → 構造化 → 出力
'============================================================
Option Explicit

Public Sub ParseCowProfile(ByVal doc As Object, ByRef kotai As Variant)
    ' 個体情報（生年月日など）を配列/UDTへ
End Sub

Public Sub ParseMovements(ByVal doc As Object, ByRef idou As Variant)
    ' 異動一覧を配列/UDTへ
End Sub

Public Sub OutputRecord(ByVal kotai As Variant, ByVal idou As Variant, ByVal farmName As String, ByVal initialGroup As Integer)
    ' 「詳細」「集計」シートに1頭分を書き出す（既存 output を移植）
End Sub

