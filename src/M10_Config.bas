Attribute VB_Name = "M10_Config"
'============================================================
' M10_Config : 定数/設定・期首/期末日付の取得
'============================================================
Option Explicit

Public Const SHEET_INPUT As String = "入力シート"

Public Sub PrepareSheetsAndDates()
    Dim ws As Worksheet, wb As Workbook: Set wb = ThisWorkbook

    ' ① 入力シート以外削除
    Application.DisplayAlerts = False
    For Each ws In wb.Worksheets
        If ws.Name <> SHEET_INPUT Then ws.Delete
    Next ws
    Application.DisplayAlerts = True

    ' ② end_date
    Dim y As Variant, m As Variant, d As Variant
    With wb.Sheets(SHEET_INPUT)
        y = .Range("H2").value: m = .Range("I2").value: d = .Range("J2").value
    End With
    If IsNumeric(y) And IsNumeric(m) And IsNumeric(d) Then
        end_date = DateSerial(CLng(y), CLng(m), CLng(d))
    Else
        MsgBox "決算年月日（H2:I2:J2）の入力が不足しています。", vbExclamation
        End
    End If

    ' ③ start_date（なければ1年前を自動設定）
    Dim sy As Variant, sm As Variant, sd As Variant
    With wb.Sheets(SHEET_INPUT)
        sy = .Range("H6").value: sm = .Range("I6").value: sd = .Range("J6").value
        If IsEmpty(sy) And IsEmpty(sm) And IsEmpty(sd) Then
            start_date = DateAdd("yyyy", -1, end_date) + 1
        Else
            start_date = DateSerial(CLng(sy), CLng(sm), CLng(sd))
        End If
    End With
End Sub

