Attribute VB_Name = "M00_AppEntry"
Option Explicit

Public farm_list As Variant
Public farm_name As String
Public start_date As Date
Public end_date As Date

Sub MAIN()
    Dim start_time As Date, end_time As Date

    '■実行確認
    If Not ConfirmExecution Then Exit Sub

    start_time = Time

    '■前処理：シート初期化・期首／期末日付取得
    Call PrepareSheetsAndDates

    '■データ取得（期首／期末／異動牛）
    Dim start_cows As Variant, end_cows As Variant, moved_cows As Variant
    Call LoadInitialCowLists(start_cows, end_cows, moved_cows)

    '■出力シートの作成（詳細・集計）
    Call CreateOutputSheets

    '■牛ID統合・重複除去
    Dim cow_list As Variant
    cow_list = BuildUniqueCowList(start_cows, end_cows, moved_cows)
    If IsEmpty(cow_list) Or end_date = "1999/12/31" Then
        MsgBox "入力データがありません。終了します。"
        Exit Sub
    End If

    '■牧場名の取得（最初の1頭を使って識別）
    farm_name = GetFarmName(cow_list)

    '■期首・期末に記録があった牛をスクレイピング
    Call SearchInitialCows(cow_list)

    '■当期出生分（推定）の牛を検索
    Call SearchGeneratedCows

    '■集計シート整形・計算式入力
    Call FormatOutputSheets

    '■終了通知と処理時間表示
    end_time = Time
    Call NotifyCompletion(start_time, end_time)
End Sub


'------------------------------------------------------
' ユーザーにマクロ実行可否を確認する
' 実行する場合は True、キャンセルの場合は False を返す
'------------------------------------------------------
Function ConfirmExecution() As Boolean
    Dim userResponse As VbMsgBoxResult
    
    userResponse = MsgBox( _
        "マクロを実行しますか？" & vbCrLf & _
        "（想定処理時間: 20?30分）", _
        vbYesNo + vbQuestion, _
        "実行確認")
    
    If userResponse = vbYes Then
        ConfirmExecution = True
    Else
        ConfirmExecution = False
    End If
End Function




