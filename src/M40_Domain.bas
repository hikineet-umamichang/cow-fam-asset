Attribute VB_Name = "M40_Domain"
'============================================================
' M40_Domain
' 目的 : 牛ID生成・チェックディジット・業務ロジック
' 依存 : なし（下位のUTLのみ）
' 提供 : GenerateCowID, AddCheckDigit, ClassifyMovement ...
'============================================================
Option Explicit


' 指定された個体識別番号をintervalだけ増減させ、新しいチェックディジット付き番号を返す
Function GenerateCowID(ByVal cowIDNumber As String, ByVal interval As Long) As String
    Dim farmCode As Long
    Dim serialNumber As Long
    Dim newSerialBase As Long
    Dim baseWithoutCheckDigit As Long

    ' 先頭5桁 = 牧場コード、次の4桁 = 連番
    farmCode = CLng(Left(cowIDNumber, 5))
    serialNumber = CLng(Mid(cowIDNumber, 6, 4))

    ' 連番をintervalだけ加減算
    newSerialBase = serialNumber + interval
    
    ' 4桁でロールオーバー（0000 → 9999）する場合の補正（※必要なら）
    If newSerialBase > 9999 Then newSerialBase = newSerialBase - 10000
    If newSerialBase < 0 Then newSerialBase = newSerialBase + 10000

    ' チェックディジットを除いた9桁の数字を作成
    baseWithoutCheckDigit = farmCode * 10000 + newSerialBase

    ' チェックディジットを付けて返す
    GenerateCowID = AddCheckDigit(baseWithoutCheckDigit)
End Function


' チェックディジットを付けた個体識別番号（10桁）を返す
Function AddCheckDigit(ByVal number As Long) As String
    Dim evenSum As Long, oddSum As Long
    Dim numberStr As String
    Dim i As Integer

    numberStr = Format(number, "000000000") ' 9桁に0埋め

    ' 奇数桁と偶数桁で分けて加算（1-based index）
    For i = 1 To 9
        If i Mod 2 = 0 Then
            evenSum = evenSum + CInt(Mid(numberStr, i, 1))
        Else
            oddSum = oddSum + CInt(Mid(numberStr, i, 1))
        End If
    Next i

    ' チェックディジット計算（モジュラス10補数）
    Dim checkDigit As Integer
    checkDigit = (10 - ((oddSum * 3 + evenSum) Mod 10)) Mod 10

    AddCheckDigit = numberStr & CStr(checkDigit)
End Function

