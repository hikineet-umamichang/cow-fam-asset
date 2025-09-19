Attribute VB_Name = "M30_Scrape"
'============================================================
' M30_Scrape : Selenium起動/同意～検索/結果取得（堅牢化版）
'============================================================
Option Explicit

' --- 定数（必要に応じて M10_Config に移してもOK） ---
Private Const BASE_URL As String = "https://www.id.nlbc.go.jp/CattleSearch/search/agreement"
Private Const CONSENT_BUTTON_NAME As String = "method:goSearch"
Private Const INPUT_NAME_IDNO As String = "txtIDNO"
Private Const RESULT_READY_ID As String = "print"        ' 結果画面の存在確認に使用
Private Const MAX_RETRY_NAV As Long = 5                  ' 画面遷移の最大リトライ
Private Const MAX_RETRY_READY As Long = 5                ' 結果待ちの最大ポーリング
Private Const BASE_WAIT_MS As Long = 120                 ' バックオフの基準待機(ms)

'============================================================
' Driver 初期化
'============================================================
Public Function InitDriver() As Selenium.ChromeDriver
    Dim drv As New Selenium.ChromeDriver
    drv.AddArgument "headless"
    drv.AddArgument "disable-gpu"
    SafeOpen drv, Chrome          ' 既存のSafeOpenを利用
    Set InitDriver = drv
End Function

'============================================================
' 個体IDで検索して、個体情報/異動情報を取得
' 成功: True（kotai/idou に2次元配列）、失敗: False（kotai/idou は Empty）
'============================================================
Public Function CowSearch(ByVal drv As Selenium.ChromeDriver, _
                          ByVal cowID As String, _
                          ByRef kotai As Variant, _
                          ByRef idou As Variant) As Boolean
    Dim ok As Boolean
    kotai = Empty
    idou = Empty

    ' 1) 検索画面へ（同意ページ→検索ページ）
    ok = NavigateToSearchPage(drv)
    If Not ok Then Exit Function

    ' 2) ID入力＆検索
    ok = InputAndSubmitID(drv, cowID)
    If Not ok Then Exit Function

    ' 3) 結果待ち→テーブル読み出し
    ok = TryGetResultTables(drv, kotai, idou)
    CowSearch = ok
End Function

'============================================================
' 同意ページにアクセス→同意ボタン押下→検索画面へ
'============================================================
Private Function NavigateToSearchPage(ByVal drv As Selenium.ChromeDriver) As Boolean
    Dim myBy As New By
    Dim retry As Long

    For retry = 0 To MAX_RETRY_NAV
        On Error Resume Next
        drv.Get BASE_URL
        On Error GoTo 0

        ' 同意ボタンが現れるまで待機（指数バックオフ込み）
        If WaitUntilPresent(drv, myBy.Name(CONSENT_BUTTON_NAME), 3000) Then
            drv.FindElementByName(CONSENT_BUTTON_NAME).Click
            NavigateToSearchPage = True
            Exit Function
        End If

        ' リトライ
        BackoffWait retry
    Next retry

    NavigateToSearchPage = False
End Function

'============================================================
' ID入力→Enter送信
'============================================================
Private Function InputAndSubmitID(ByVal drv As Selenium.ChromeDriver, ByVal cowID As String) As Boolean
    Dim myBy As New By
    Dim sKey As New Selenium.keys

    ' 入力欄待ち
    If Not WaitUntilPresent(drv, myBy.Name(INPUT_NAME_IDNO), 3000) Then
        InputAndSubmitID = False
        Exit Function
    End If

    ' 10桁ゼロパディング
    Dim id10 As String
    id10 = Format$(cowID, "0000000000")

    drv.FindElementByName(INPUT_NAME_IDNO).Clear
    drv.FindElementByName(INPUT_NAME_IDNO).SendKeys id10
    drv.FindElementByName(INPUT_NAME_IDNO).SendKeys sKey.Enter

    InputAndSubmitID = True
End Function

'============================================================
' 結果画面の準備完了を待ち、個体/異動テーブルを取り出す
'============================================================
Private Function TryGetResultTables(ByVal drv As Selenium.ChromeDriver, _
                                    ByRef kotai As Variant, _
                                    ByRef idou As Variant) As Boolean
    Dim myBy As New By
    Dim retry As Long

    ' 結果画面のアンカー要素（id="print"）が出るまで待つ
    For retry = 0 To MAX_RETRY_READY
        If drv.IsElementPresent(myBy.ID(RESULT_READY_ID)) Then Exit For
        BackoffWait retry
    Next retry

    If Not drv.IsElementPresent(myBy.ID(RESULT_READY_ID)) Then
        ' 見つからない＝該当なし/タイムアウト
        kotai = Empty
        idou = Empty
        TryGetResultTables = False
        Exit Function
    End If

    ' 画面内のテーブルから既存インデックス(8,9番目)相当を安全に取得
    Dim tbls As Selenium.WebElements
    Set tbls = drv.FindElementsByTag("Table")

    If tbls Is Nothing Or tbls.Count < 9 Then
        ' 期待するテーブルに届かない場合は失敗として扱う
        kotai = Empty
        idou = Empty
        TryGetResultTables = False
        Exit Function
    End If

    On Error Resume Next
    kotai = tbls.Item(7).AsTable.Data   ' 0-basedのため7=8番目
    idou = tbls.Item(8).AsTable.Data    ' 8=9番目
    On Error GoTo 0

    ' どちらか空なら失敗扱い
    If IsEmpty(kotai) And IsEmpty(idou) Then
        TryGetResultTables = False
    Else
        TryGetResultTables = True
    End If
End Function

'============================================================
' 要素の存在をポーリングして待つ（timeoutMs まで）
'============================================================
Private Function WaitUntilPresent(ByVal drv As Selenium.ChromeDriver, _
                                  ByVal locator As By, _
                                  ByVal timeoutMs As Long, _
                                  Optional ByVal pollMs As Long = 150) As Boolean
    Dim tStart As Single: tStart = Timer
    Do
        If drv.IsElementPresent(locator) Then
            WaitUntilPresent = True
            Exit Function
        End If
        drv.Wait pollMs
        If (Timer - tStart) * 1000# >= timeoutMs Then Exit Do
    Loop
    WaitUntilPresent = False
End Function

'============================================================
' 指数バックオフ待機（retry回数に応じて待機を伸ばす＋ランダマイズ）
'============================================================
Private Sub BackoffWait(ByVal retry As Long)
    Randomize
    Dim ms As Long
    ms = CLng(BASE_WAIT_MS * (2 ^ retry) + Rnd() * 100#)
    If ms > 3000 Then ms = 3000   ' 上限キャップ（必要に応じて調整）
    Application.Wait Now + TimeSerial(0, 0, 0) ' tick flush
    DoEvents
    ' SeleniumBasicのWaitはドライバ側の待機
    Dim drvWait As New Selenium.ChromeDriver ' ダミーを作らない
    ' 代替としてVBA側の精度ある待機
    Dim t As Single: t = Timer
    Do While (Timer - t) * 1000# < ms
        DoEvents
    Loop
End Sub
