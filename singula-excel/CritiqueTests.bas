Attribute VB_Name = "CritiqueTests"
Option Explicit

' Run manually inside Windows Excel. Makes no network requests.
Public Sub RunCritiqueSelfTests()
    Dim parsed As Object, fixture As String, sample As String, count As Long
    On Error GoTo Failed
    sample = "日本語の批評" & vbCrLf & Chr$(34) & "引用" & Chr$(34) & Chr$(92) & vbTab & ChrW$(-10179) & ChrW$(-8704)
    Set parsed = JsonParse("{""text"":" & JsonQuote(sample) & "}")
    Check parsed("text") = sample, "JSON Japanese, escaping and surrogate pair round trip", count
    Set parsed = JsonParse("{""text"":""\u65e5\u672c\u8a9e\ud83d\ude00"",""array"":[true,false,null,-1.25e2,{""x"":1}]}")
    Check parsed("text") = "日本語" & ChrW$(-10179) & ChrW$(-8704), "JSON Unicode escape decoding", count
    Check parsed("array").Count = 5, "JSON array length", count
    Check parsed("array")(1) = True, "JSON boolean", count
    Check IsNull(parsed("array")(3)), "JSON null", count
    Check parsed("array")(4) = -125, "JSON number", count
    Check parsed("array")(5)("x") = 1, "JSON nested object", count
    Check InvalidJson("{""x"":01}"), "Reject leading zero", count
    Check InvalidJson("{""x"":""\q""}"), "Reject invalid escape", count
    Check InvalidJson("{""x"":1,}"), "Reject trailing comma", count
    Check InvalidJson("{""x"":1} extra"), "Reject trailing data", count
    Check InvalidJson("{""x"":1,""x"":2}"), "Reject duplicate key", count
    fixture = "{""status"":""completed"",""output"":[{""type"":""reasoning""},{""type"":""message"",""role"":""assistant"",""content"":[{""type"":""output_text"",""text"":" & JsonQuote("批評の文章") & "}]}]}"
    Check ParseCritiqueResponse(fixture) = "批評の文章", "Responses API output_text extraction", count
    Check InvalidResponse("{""status"":""incomplete"",""output"":[]}"), "Reject incomplete generation", count
    Check InvalidResponse("{""status"":""completed"",""output"":[]}"), "Reject empty generation", count
    Check InvalidResponse("{""output"":[]}"), "Reject missing status", count
    Check InvalidResponse("{""status"":""completed"",""output"":[{""type"":""message"",""role"":""assistant"",""content"":[{""type"":""refusal"",""refusal"":""No""}]}]}"), "Reject refusal", count
    MsgBox CStr(count) & "件のオフラインテストが成功しました。API通信・実ファイル出力の確認は別途必要です。", vbInformation
    Exit Sub
Failed:
    MsgBox "テスト失敗: " & Err.Description, vbCritical
End Sub

' Integration tests for the offline Excel runner, using a private temporary folder.
Public Sub RunExcelMockSelfTests()
    Dim fso As Object, folder As String, outputPath As String, stage As Long, failureStage As Long
    Dim reviews(1 To 5) As String, overall As String, summary As String, result As String
    Dim state As Object, count As Long, failureMessage As String, text As String, hashBytes() As Byte
    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = fso.BuildPath(fso.GetSpecialFolder(2), fso.GetTempName)
    fso.CreateFolder folder
    outputPath = fso.BuildPath(folder, "mock-result.txt")
    hashBytes = StrConv("abc", vbFromUnicode)
    Check MockSha256(hashBytes) = "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad", "SHA-256 abc vector", count
    hashBytes = StrConv("abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq", vbFromUnicode)
    Check MockSha256(hashBytes) = "248d6a61d20638b8e5c026930c3e6039a33ce45964ff2167f6ecedd419db06c1", "SHA-256 multi-block vector", count
    Check MockFingerprint("sample.txt", "abc") = "6cf6dc128ec721c696949294f4e3bd39f7ef0e2687cfc344b219eea15dd32eb0", "SHA-256 known value", count
    BeginExcelMock "本文識別子", "sample.txt", outputPath, False
    reviews(1) = ExcelMockReview(1, "独立評価", "本文識別子", 0, 0)
    reviews(2) = ExcelMockReview(2, "独立評価", "本文識別子", 0, 0)
    Check MockStepError(3, 3, 0) = EXCEL_MOCK_STOP, "Stop after stage three", count
    Check Not fso.FileExists(outputPath), "No complete output while stopped", count
    Set state = JsonParse(NovelCritique.ReadText(outputPath & ".excel-mock-state.json", "utf-8"))
    Check state("results").Count = 3, "Three results retained", count
    Check InStr(1, NovelCritique.ReadText(outputPath & ".excel-mock-state.json", "utf-8"), "本文識別子") = 0, "Novel excluded from checkpoint", count
    BeginExcelMock "本文識別子", "sample.txt", outputPath, True
    Check ExcelMockReview(1, "独立評価", "本文識別子", 0, 1) = reviews(1), "Cached result bypasses failure injection", count
    For stage = 1 To 5
        reviews(stage) = ExcelMockReview(stage, "独立評価", "本文識別子", 0, 0)
        summary = summary & reviews(stage) & vbCrLf
    Next
    overall = ExcelMockReview(6, "批評のみから総評", summary, 0, 0)
    result = FormatCritiqueResult("sample.txt", reviews, overall)
    Check InStr(result, "審査員１" & vbCrLf & EXCEL_MOCK_MARKER) > 0, "Review layout and marker", count
    Check InStr(result, "総評" & vbCrLf & EXCEL_MOCK_MARKER) > 0, "Summary layout and marker", count
    NovelCritique.WriteResult outputPath, result
    Check NovelCritique.ReadText(outputPath, "utf-8") = result, "UTF-8 file output round trip", count
    Check SaveResultError(outputPath, "overwrite") <> 0, "Existing output rejected", count
    Check NovelCritique.ReadText(outputPath, "utf-8") = result, "Existing output preserved", count
    Check BeginMockError("変更後の本文", "sample.txt", outputPath, True) <> 0, "Changed novel rejected", count
    For failureStage = 1 To 6
        outputPath = fso.BuildPath(folder, "error-" & CStr(failureStage) & ".txt")
        BeginExcelMock "本文識別子", "sample.txt", outputPath, False
        For stage = 1 To failureStage - 1
            text = ExcelMockReview(stage, "test", "test", 0, 0)
        Next
        Check MockStepError(failureStage, 0, failureStage) = EXCEL_MOCK_FAILURE, "Injected error " & CStr(failureStage), count
        Check Not fso.FileExists(outputPath), "No completed result at error " & CStr(failureStage), count
        BeginExcelMock "本文識別子", "sample.txt", outputPath, True
        For stage = 1 To 6
            text = ExcelMockReview(stage, "test", "test", 0, 0)
        Next
        Set state = JsonParse(NovelCritique.ReadText(outputPath & ".excel-mock-state.json", "utf-8"))
        Check state("results").Count = 6, "Resume after error " & CStr(failureStage), count
    Next
    outputPath = fso.BuildPath(folder, "stop-six.txt")
    BeginExcelMock "本文識別子", "sample.txt", outputPath, False
    For stage = 1 To 5
        text = ExcelMockReview(stage, "test", "test", 0, 0)
    Next
    Check MockStepError(6, 6, 0) = EXCEL_MOCK_STOP, "Stop after summary", count
    BeginExcelMock "本文識別子", "sample.txt", outputPath, True
    Check Left$(ExcelMockReview(6, "test", "test", 0, 6), Len(EXCEL_MOCK_MARKER)) = EXCEL_MOCK_MARKER, "Resume saved summary", count
    fso.DeleteFolder folder, True
    MsgBox CStr(count) & "件のExcelモックテストが成功しました。通信は行っていません。", vbInformation
    Exit Sub
Failed:
    failureMessage = Err.Description
    On Error Resume Next
    If Len(folder) > 0 Then fso.DeleteFolder folder, True
    On Error GoTo 0
    MsgBox "Excelモックテスト失敗: " & failureMessage, vbCritical
End Sub

Private Function MockStepError(ByVal stage As Long, ByVal stopAfter As Long, ByVal failAt As Long) As Long
    Dim text As String
    On Error GoTo Rejected
    text = ExcelMockReview(stage, "test", "test", stopAfter, failAt)
    Exit Function
Rejected:
    MockStepError = Err.Number
End Function

Private Function BeginMockError(ByVal novel As String, ByVal filename As String, ByVal path As String, ByVal resumeMock As Boolean) As Long
    On Error GoTo Rejected
    BeginExcelMock novel, filename, path, resumeMock
    Exit Function
Rejected:
    BeginMockError = Err.Number
End Function

Private Function SaveResultError(ByVal path As String, ByVal text As String) As Long
    On Error GoTo Rejected
    NovelCritique.WriteResult path, text
    Exit Function
Rejected:
    SaveResultError = Err.Number
End Function

Public Sub RunExcelThresholdSelfTests()
    Dim count As Long, unlimited As Variant, text As String
    On Error GoTo Failed
    text = "あ" & vbCrLf & " " & ChrW$(-10179) & ChrW$(-8704)
    Check NovelCharacterCount(text) = 4, "Unicode code points, CRLF and spaces", count
    Check NovelCharacterCount("a" & vbCr & "b") = 3, "CR normalized to one newline", count
    Check NovelCharacterCount("") = 0, "Empty count", count
    Check IsEmpty(ParseCharacterLimit("")), "Blank means unlimited", count
    Check ParseCharacterLimit(0) = 0, "Zero is a valid threshold", count
    Check ExcelThresholdResult("sample.txt", "abc", 4, unlimited) = "sample.txt" & vbCrLf & "閾値の下限を下回っているため評価無し" & vbCrLf, "Lower threshold exact output", count
    Check ExcelThresholdResult("sample.txt", "abc", unlimited, 2) = "sample.txt" & vbCrLf & "閾値の上限を上回っているので評価無し" & vbCrLf, "Upper threshold exact output", count
    Check ExcelThresholdResult("sample.txt", "abc", 3, 3) = "", "Inclusive boundaries", count
    Check ExcelThresholdResult("sample.txt", "abc", unlimited, unlimited) = "", "Unlimited", count
    Check ExcelThresholdResult("sample.txt", "", 1, unlimited) <> "", "Empty input below lower bound", count
    Check ExcelThresholdResult("sample.txt", text, 4, 4) = "", "Unicode boundary", count
    Check InStr(ExcelThresholdResult("sample.txt", "abc", 4, unlimited, True), EXCEL_MOCK_MARKER) > 0, "Mock skip marker", count
    Check ThresholdError(-1, unlimited) <> 0, "Reject negative lower", count
    Check ThresholdError(unlimited, -1) <> 0, "Reject negative upper", count
    Check ThresholdError(1.5, unlimited) <> 0, "Reject fractional value", count
    Check ThresholdError("invalid", unlimited) <> 0, "Reject text", count
    Check ThresholdError(4, 3) <> 0, "Reject reversed range", count
    Check ThresholdError(2147483648#, unlimited) <> 0, "Reject Long overflow", count
    Check ThresholdError(True, unlimited) <> 0, "Reject Boolean", count
    MsgBox CStr(count) & "件の文字数閾値テストが成功しました。通信は行っていません。", vbInformation
    Exit Sub
Failed:
    MsgBox "文字数閾値テスト失敗: " & Err.Description, vbCritical
End Sub

Private Function ThresholdError(ByVal minChars As Variant, ByVal maxChars As Variant) As Long
    Dim result As String
    On Error GoTo Rejected
    result = ExcelThresholdResult("sample.txt", "abc", minChars, maxChars)
    Exit Function
Rejected:
    ThresholdError = Err.Number
End Function

Private Sub Check(ByVal condition As Boolean, ByVal label As String, ByRef count As Long)
    If Not condition Then Err.Raise vbObjectError + 2300, , label
    count = count + 1
End Sub

Private Function InvalidJson(ByVal text As String) As Boolean
    Dim result As Object
    On Error GoTo Rejected
    Set result = JsonParse(text)
    Exit Function
Rejected:
    InvalidJson = True
End Function

Private Function InvalidResponse(ByVal text As String) As Boolean
    Dim result As String
    On Error GoTo Rejected
    result = ParseCritiqueResponse(text)
    Exit Function
Rejected:
    InvalidResponse = True
End Function
