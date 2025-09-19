Attribute VB_Name = "M30_Scrape"
'============================================================
' M30_Scrape : Selenium�N��/���Ӂ`����/���ʎ擾�i���S���Łj
'============================================================
Option Explicit

' --- �萔�i�K�v�ɉ����� M10_Config �Ɉڂ��Ă�OK�j ---
Private Const BASE_URL As String = "https://www.id.nlbc.go.jp/CattleSearch/search/agreement"
Private Const CONSENT_BUTTON_NAME As String = "method:goSearch"
Private Const INPUT_NAME_IDNO As String = "txtIDNO"
Private Const RESULT_READY_ID As String = "print"        ' ���ʉ�ʂ̑��݊m�F�Ɏg�p
Private Const MAX_RETRY_NAV As Long = 5                  ' ��ʑJ�ڂ̍ő僊�g���C
Private Const MAX_RETRY_READY As Long = 5                ' ���ʑ҂��̍ő�|�[�����O
Private Const BASE_WAIT_MS As Long = 120                 ' �o�b�N�I�t�̊�ҋ@(ms)

'============================================================
' Driver ������
'============================================================
Public Function InitDriver() As Selenium.ChromeDriver
    Dim drv As New Selenium.ChromeDriver
    drv.AddArgument "headless"
    drv.AddArgument "disable-gpu"
    SafeOpen drv, Chrome          ' ������SafeOpen�𗘗p
    Set InitDriver = drv
End Function

'============================================================
' ��ID�Ō������āA�̏��/�ٓ������擾
' ����: True�ikotai/idou ��2�����z��j�A���s: False�ikotai/idou �� Empty�j
'============================================================
Public Function CowSearch(ByVal drv As Selenium.ChromeDriver, _
                          ByVal cowID As String, _
                          ByRef kotai As Variant, _
                          ByRef idou As Variant) As Boolean
    Dim ok As Boolean
    kotai = Empty
    idou = Empty

    ' 1) ������ʂցi���Ӄy�[�W�������y�[�W�j
    ok = NavigateToSearchPage(drv)
    If Not ok Then Exit Function

    ' 2) ID���́�����
    ok = InputAndSubmitID(drv, cowID)
    If Not ok Then Exit Function

    ' 3) ���ʑ҂����e�[�u���ǂݏo��
    ok = TryGetResultTables(drv, kotai, idou)
    CowSearch = ok
End Function

'============================================================
' ���Ӄy�[�W�ɃA�N�Z�X�����Ӄ{�^��������������ʂ�
'============================================================
Private Function NavigateToSearchPage(ByVal drv As Selenium.ChromeDriver) As Boolean
    Dim myBy As New By
    Dim retry As Long

    For retry = 0 To MAX_RETRY_NAV
        On Error Resume Next
        drv.Get BASE_URL
        On Error GoTo 0

        ' ���Ӄ{�^���������܂őҋ@�i�w���o�b�N�I�t���݁j
        If WaitUntilPresent(drv, myBy.Name(CONSENT_BUTTON_NAME), 3000) Then
            drv.FindElementByName(CONSENT_BUTTON_NAME).Click
            NavigateToSearchPage = True
            Exit Function
        End If

        ' ���g���C
        BackoffWait retry
    Next retry

    NavigateToSearchPage = False
End Function

'============================================================
' ID���́�Enter���M
'============================================================
Private Function InputAndSubmitID(ByVal drv As Selenium.ChromeDriver, ByVal cowID As String) As Boolean
    Dim myBy As New By
    Dim sKey As New Selenium.keys

    ' ���͗��҂�
    If Not WaitUntilPresent(drv, myBy.Name(INPUT_NAME_IDNO), 3000) Then
        InputAndSubmitID = False
        Exit Function
    End If

    ' 10���[���p�f�B���O
    Dim id10 As String
    id10 = Format$(cowID, "0000000000")

    drv.FindElementByName(INPUT_NAME_IDNO).Clear
    drv.FindElementByName(INPUT_NAME_IDNO).SendKeys id10
    drv.FindElementByName(INPUT_NAME_IDNO).SendKeys sKey.Enter

    InputAndSubmitID = True
End Function

'============================================================
' ���ʉ�ʂ̏���������҂��A��/�ٓ��e�[�u�������o��
'============================================================
Private Function TryGetResultTables(ByVal drv As Selenium.ChromeDriver, _
                                    ByRef kotai As Variant, _
                                    ByRef idou As Variant) As Boolean
    Dim myBy As New By
    Dim retry As Long

    ' ���ʉ�ʂ̃A���J�[�v�f�iid="print"�j���o��܂ő҂�
    For retry = 0 To MAX_RETRY_READY
        If drv.IsElementPresent(myBy.ID(RESULT_READY_ID)) Then Exit For
        BackoffWait retry
    Next retry

    If Not drv.IsElementPresent(myBy.ID(RESULT_READY_ID)) Then
        ' ������Ȃ����Y���Ȃ�/�^�C���A�E�g
        kotai = Empty
        idou = Empty
        TryGetResultTables = False
        Exit Function
    End If

    ' ��ʓ��̃e�[�u����������C���f�b�N�X(8,9�Ԗ�)���������S�Ɏ擾
    Dim tbls As Selenium.WebElements
    Set tbls = drv.FindElementsByTag("Table")

    If tbls Is Nothing Or tbls.Count < 9 Then
        ' ���҂���e�[�u���ɓ͂��Ȃ��ꍇ�͎��s�Ƃ��Ĉ���
        kotai = Empty
        idou = Empty
        TryGetResultTables = False
        Exit Function
    End If

    On Error Resume Next
    kotai = tbls.Item(7).AsTable.Data   ' 0-based�̂���7=8�Ԗ�
    idou = tbls.Item(8).AsTable.Data    ' 8=9�Ԗ�
    On Error GoTo 0

    ' �ǂ��炩��Ȃ玸�s����
    If IsEmpty(kotai) And IsEmpty(idou) Then
        TryGetResultTables = False
    Else
        TryGetResultTables = True
    End If
End Function

'============================================================
' �v�f�̑��݂��|�[�����O���đ҂itimeoutMs �܂Łj
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
' �w���o�b�N�I�t�ҋ@�iretry�񐔂ɉ����đҋ@��L�΂��{�����_�}�C�Y�j
'============================================================
Private Sub BackoffWait(ByVal retry As Long)
    Randomize
    Dim ms As Long
    ms = CLng(BASE_WAIT_MS * (2 ^ retry) + Rnd() * 100#)
    If ms > 3000 Then ms = 3000   ' ����L���b�v�i�K�v�ɉ����Ē����j
    Application.Wait Now + TimeSerial(0, 0, 0) ' tick flush
    DoEvents
    ' SeleniumBasic��Wait�̓h���C�o���̑ҋ@
    Dim drvWait As New Selenium.ChromeDriver ' �_�~�[�����Ȃ�
    ' ��ւƂ���VBA���̐��x����ҋ@
    Dim t As Single: t = Timer
    Do While (Timer - t) * 1000# < ms
        DoEvents
    Loop
End Sub
