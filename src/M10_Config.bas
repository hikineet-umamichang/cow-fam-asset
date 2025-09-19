Attribute VB_Name = "M10_Config"
'------------------------------------------------------
' シートを整理し、期首日付と決算日付を取得する処理
' ・"入力シート" 以外を削除
' ・決算年月日 (H2, I2, J2) → end_date
' ・期首年月日 (H6, I6, J6) → start_date
'   入力がない場合は、end_dateの1年前を自動設定
'------------------------------------------------------
Sub PrepareSheetsAndDates()
    Dim ws As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    ' ① 入力シート以外を削除
    Application.DisplayAlerts = False
    For Each ws In wb.Worksheets
        If ws.Name <> "入力シート" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
    
    ' ② 決算年月日を取得
    Dim yearVal As Variant, monthVal As Variant, dayVal As Variant
    With wb.Sheets("入力シート")
        yearVal = .Range("H2").Value
        monthVal = .Range("I2").Value
        dayVal = .Range("J2").Value
    End With
    
    If IsNumeric(yearVal) And IsNumeric(monthVal) And IsNumeric(dayVal) Then
        end_date = DateSerial(CLng(yearVal), CLng(monthVal), CLng(dayVal))
    Else
        MsgBox "決算年月日の入力が不足しています（H2:I2:J2）。処理を終了します。", vbExclamation
        End
    End If
    
    ' ③ 期首年月日を取得
    Dim startYear As Variant, startMonth As Variant, startDay As Variant
    With wb.Sheets("入力シート")
        startYear = .Range("H6").Value
        startMonth = .Range("I6").Value
        startDay = .Range("J6").Value
    End With
    
    If IsNumeric(startYear) And IsNumeric(startMonth) And IsNumeric(startDay) Then
        start_date = DateSerial(CLng(startYear), CLng(startMonth), CLng(startDay))
    Else
        ' 入力がない場合は決算日の1年前
        start_date = DateAdd("yyyy", -1, end_date)
        wb.Sheets("入力シート").Range("H6").Value = Year(start_date)
        wb.Sheets("入力シート").Range("I6").Value = Month(start_date)
        wb.Sheets("入力シート").Range("J6").Value = Day(start_date)
    End If
End Sub

