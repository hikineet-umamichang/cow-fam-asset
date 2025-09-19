Attribute VB_Name = "M10_Config"
'------------------------------------------------------
' �V�[�g�𐮗����A������t�ƌ��Z���t���擾���鏈��
' �E"���̓V�[�g" �ȊO���폜
' �E���Z�N���� (H2, I2, J2) �� end_date
' �E����N���� (H6, I6, J6) �� start_date
'   ���͂��Ȃ��ꍇ�́Aend_date��1�N�O�������ݒ�
'------------------------------------------------------
Sub PrepareSheetsAndDates()
    Dim ws As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    ' �@ ���̓V�[�g�ȊO���폜
    Application.DisplayAlerts = False
    For Each ws In wb.Worksheets
        If ws.Name <> "���̓V�[�g" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
    
    ' �A ���Z�N�������擾
    Dim yearVal As Variant, monthVal As Variant, dayVal As Variant
    With wb.Sheets("���̓V�[�g")
        yearVal = .Range("H2").Value
        monthVal = .Range("I2").Value
        dayVal = .Range("J2").Value
    End With
    
    If IsNumeric(yearVal) And IsNumeric(monthVal) And IsNumeric(dayVal) Then
        end_date = DateSerial(CLng(yearVal), CLng(monthVal), CLng(dayVal))
    Else
        MsgBox "���Z�N�����̓��͂��s�����Ă��܂��iH2:I2:J2�j�B�������I�����܂��B", vbExclamation
        End
    End If
    
    ' �B ����N�������擾
    Dim startYear As Variant, startMonth As Variant, startDay As Variant
    With wb.Sheets("���̓V�[�g")
        startYear = .Range("H6").Value
        startMonth = .Range("I6").Value
        startDay = .Range("J6").Value
    End With
    
    If IsNumeric(startYear) And IsNumeric(startMonth) And IsNumeric(startDay) Then
        start_date = DateSerial(CLng(startYear), CLng(startMonth), CLng(startDay))
    Else
        ' ���͂��Ȃ��ꍇ�͌��Z����1�N�O
        start_date = DateAdd("yyyy", -1, end_date)
        wb.Sheets("���̓V�[�g").Range("H6").Value = Year(start_date)
        wb.Sheets("���̓V�[�g").Range("I6").Value = Month(start_date)
        wb.Sheets("���̓V�[�g").Range("J6").Value = Day(start_date)
    End If
End Sub

