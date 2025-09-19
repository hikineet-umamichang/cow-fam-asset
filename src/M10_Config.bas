Attribute VB_Name = "M10_Config"
'============================================================
' M10_Config : �萔/�ݒ�E����/�������t�̎擾
'============================================================
Option Explicit

Public Const SHEET_INPUT As String = "���̓V�[�g"

Public Sub PrepareSheetsAndDates()
    Dim ws As Worksheet, wb As Workbook: Set wb = ThisWorkbook

    ' �@ ���̓V�[�g�ȊO�폜
    Application.DisplayAlerts = False
    For Each ws In wb.Worksheets
        If ws.Name <> SHEET_INPUT Then ws.Delete
    Next ws
    Application.DisplayAlerts = True

    ' �A end_date
    Dim y As Variant, m As Variant, d As Variant
    With wb.Sheets(SHEET_INPUT)
        y = .Range("H2").value: m = .Range("I2").value: d = .Range("J2").value
    End With
    If IsNumeric(y) And IsNumeric(m) And IsNumeric(d) Then
        end_date = DateSerial(CLng(y), CLng(m), CLng(d))
    Else
        MsgBox "���Z�N�����iH2:I2:J2�j�̓��͂��s�����Ă��܂��B", vbExclamation
        End
    End If

    ' �B start_date�i�Ȃ����1�N�O�������ݒ�j
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

