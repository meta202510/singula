Attribute VB_Name = "MockReview"
Option Explicit

Public Const EXCEL_MOCK_MARKER As String = "【モックテスト専用・本番審査ではありません】"
Public Const EXCEL_MOCK_STOP As Long = vbObjectError + 2400
Public Const EXCEL_MOCK_FAILURE As Long = vbObjectError + 2401

#If VBA7 Then
Private Declare PtrSafe Function MoveFileExW Lib "kernel32.dll" (ByVal oldPath As LongPtr, ByVal newPath As LongPtr, ByVal flags As Long) As Long
#Else
Private Declare Function MoveFileExW Lib "kernel32.dll" (ByVal oldPath As Long, ByVal newPath As Long, ByVal flags As Long) As Long
#End If

Private checkpoint As String
Private fingerprint As String
Private saved(1 To 6) As String
Private completed As Long

Public Function MockStepSetting(ByVal value As Variant) As Long
    If Len(Trim$(CStr(value))) = 0 Then Exit Function
    If Not IsNumeric(value) Then GoTo Invalid
    If CDbl(value) < 1 Or CDbl(value) > 6 Or CDbl(value) <> Fix(CDbl(value)) Then GoTo Invalid
    MockStepSetting = CLng(value)
    Exit Function
Invalid:
    Err.Raise vbObjectError + 2402, , "モックの停止回・エラー回は空欄または1～6の整数を指定してください。"
End Function

Public Function MockFingerprint(ByVal filename As String, ByVal novel As String) As String
    Dim stream As Object, bytes() As Byte, failureNumber As Long, phase As String
    On Error GoTo Failed
    phase = "UTF-8変換"
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText filename & ChrW$(0) & novel
    stream.Position = 0
    stream.Type = 1
    stream.Position = 3
    bytes = stream.Read
    stream.Close
    phase = "SHA-256計算"
    MockFingerprint = MockSha256(bytes)
    Exit Function
Failed:
    failureNumber = Err.Number
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise vbObjectError + 2403, , "モックの入力確認用ハッシュを計算できませんでした。処理: " & phase & " / エラー番号: " & CStr(failureNumber)
End Function

' SHA-256 with explicit unsigned 32-bit arithmetic. No CryptoAPI or network calls.
Private Function Unsigned32(ByVal value As Long) As Double
    Unsigned32 = value
    If value < 0 Then Unsigned32 = Unsigned32 + 4294967296#
End Function

Private Function Signed32(ByVal value As Double) As Long
    value = value - Fix(value / 4294967296#) * 4294967296#
    If value >= 2147483648# Then value = value - 4294967296#
    Signed32 = CLng(value)
End Function

Private Function Add32(ParamArray values() As Variant) As Long
    Dim total As Double, i As Long
    For i = LBound(values) To UBound(values)
        total = total + Unsigned32(CLng(values(i)))
    Next
    Add32 = Signed32(total)
End Function

Private Function ShiftRight32(ByVal value As Long, ByVal bits As Long) As Long
    ShiftRight32 = CLng(Fix(Unsigned32(value) / (2# ^ bits)))
End Function

Private Function RotateRight32(ByVal value As Long, ByVal bits As Long) As Long
    Dim unsigned As Double, low As Double, divisor As Double
    unsigned = Unsigned32(value)
    divisor = 2# ^ bits
    low = unsigned - Fix(unsigned / divisor) * divisor
    RotateRight32 = ShiftRight32(value, bits) Or Signed32(low * (2# ^ (32 - bits)))
End Function

Public Function MockSha256(ByRef inputBytes() As Byte) As String
    Dim data() As Byte, length As Long, padded As Long, offset As Long, t As Long, i As Long
    Dim words(0 To 63) As Long, state(0 To 7) As Long, constants As Variant
    Dim a As Long, b As Long, c As Long, d As Long, e As Long, f As Long, g As Long, h As Long
    Dim sigma0 As Long, sigma1 As Long, choice As Long, majority As Long, temp1 As Long, temp2 As Long
    Dim bitLength As Double, tail As Double
    length = UBound(inputBytes) - LBound(inputBytes) + 1
    padded = ((length + 8) \ 64 + 1) * 64
    ReDim data(0 To padded - 1)
    For i = 0 To length - 1
        data(i) = inputBytes(LBound(inputBytes) + i)
    Next
    data(length) = 128
    bitLength = CDbl(length) * 8#
    For i = 0 To 7
        tail = bitLength - Fix(bitLength / 256#) * 256#
        data(padded - 1 - i) = CByte(tail)
        bitLength = Fix(bitLength / 256#)
    Next
    constants = Array( _
        &H428A2F98, &H71374491, &HB5C0FBCF, &HE9B5DBA5, &H3956C25B, &H59F111F1, &H923F82A4, &HAB1C5ED5, _
        &HD807AA98, &H12835B01, &H243185BE, &H550C7DC3, &H72BE5D74, &H80DEB1FE, &H9BDC06A7, &HC19BF174, _
        &HE49B69C1, &HEFBE4786, &HFC19DC6, &H240CA1CC, &H2DE92C6F, &H4A7484AA, &H5CB0A9DC, &H76F988DA, _
        &H983E5152, &HA831C66D, &HB00327C8, &HBF597FC7, &HC6E00BF3, &HD5A79147, &H6CA6351, &H14292967, _
        &H27B70A85, &H2E1B2138, &H4D2C6DFC, &H53380D13, &H650A7354, &H766A0ABB, &H81C2C92E, &H92722C85, _
        &HA2BFE8A1, &HA81A664B, &HC24B8B70, &HC76C51A3, &HD192E819, &HD6990624, &HF40E3585, &H106AA070, _
        &H19A4C116, &H1E376C08, &H2748774C, &H34B0BCB5, &H391C0CB3, &H4ED8AA4A, &H5B9CCA4F, &H682E6FF3, _
        &H748F82EE, &H78A5636F, &H84C87814, &H8CC70208, &H90BEFFFA, &HA4506CEB, &HBEF9A3F7, &HC67178F2)
    state(0) = &H6A09E667: state(1) = &HBB67AE85: state(2) = &H3C6EF372: state(3) = &HA54FF53A
    state(4) = &H510E527F: state(5) = &H9B05688C: state(6) = &H1F83D9AB: state(7) = &H5BE0CD19
    For offset = 0 To padded - 1 Step 64
        For t = 0 To 15
            i = offset + t * 4
            words(t) = Signed32(CDbl(data(i)) * 16777216# + CDbl(data(i + 1)) * 65536# + CDbl(data(i + 2)) * 256# + data(i + 3))
        Next
        For t = 16 To 63
            sigma0 = RotateRight32(words(t - 15), 7) Xor RotateRight32(words(t - 15), 18) Xor ShiftRight32(words(t - 15), 3)
            sigma1 = RotateRight32(words(t - 2), 17) Xor RotateRight32(words(t - 2), 19) Xor ShiftRight32(words(t - 2), 10)
            words(t) = Add32(words(t - 16), sigma0, words(t - 7), sigma1)
        Next
        a = state(0): b = state(1): c = state(2): d = state(3)
        e = state(4): f = state(5): g = state(6): h = state(7)
        For t = 0 To 63
            sigma1 = RotateRight32(e, 6) Xor RotateRight32(e, 11) Xor RotateRight32(e, 25)
            choice = (e And f) Xor ((Not e) And g)
            temp1 = Add32(h, sigma1, choice, CLng(constants(t)), words(t))
            sigma0 = RotateRight32(a, 2) Xor RotateRight32(a, 13) Xor RotateRight32(a, 22)
            majority = (a And b) Xor (a And c) Xor (b And c)
            temp2 = Add32(sigma0, majority)
            h = g: g = f: f = e: e = Add32(d, temp1)
            d = c: c = b: b = a: a = Add32(temp1, temp2)
        Next
        state(0) = Add32(state(0), a): state(1) = Add32(state(1), b)
        state(2) = Add32(state(2), c): state(3) = Add32(state(3), d)
        state(4) = Add32(state(4), e): state(5) = Add32(state(5), f)
        state(6) = Add32(state(6), g): state(7) = Add32(state(7), h)
    Next
    For i = 0 To 7
        MockSha256 = MockSha256 & LCase$(Right$("00000000" & Hex$(state(i)), 8))
    Next
End Function

Public Sub BeginExcelMock(ByVal novel As String, ByVal filename As String, ByVal outputPath As String, ByVal resumeMock As Boolean)
    Dim fso As Object, state As Object, items As Object, i As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    checkpoint = outputPath & ".excel-mock-state.json"
    fingerprint = MockFingerprint(filename, novel)
    completed = 0
    Erase saved
    If resumeMock Then
        On Error GoTo InvalidState
        Set state = JsonParse(NovelCritique.ReadText(checkpoint, "utf-8"))
        If Not state.Exists("mode") Then GoTo InvalidState
        If state("mode") <> "singula-excel-mock-v1" Then GoTo InvalidState
        If Not state.Exists("fingerprint") Then GoTo InvalidState
        If state("fingerprint") <> fingerprint Then GoTo InvalidState
        If Not state.Exists("results") Then GoTo InvalidState
        If TypeName(state("results")) <> "Collection" Then GoTo InvalidState
        Set items = state("results")
        If items.Count > 6 Then GoTo InvalidState
        For i = 1 To items.Count
            If VarType(items(i)) <> vbString Then GoTo InvalidState
            If Left$(items(i), Len(EXCEL_MOCK_MARKER)) <> EXCEL_MOCK_MARKER Then GoTo InvalidState
            saved(i) = items(i)
        Next
        completed = items.Count
        On Error GoTo 0
    Else
        If fso.FileExists(checkpoint) Or fso.FolderExists(checkpoint) Then
            Err.Raise vbObjectError + 2404, , "モック途中データが存在します。再開を選ぶか別の保存先を指定してください。"
        End If
        SaveCheckpoint
    End If
    Exit Sub
InvalidState:
    On Error GoTo 0
    completed = 0
    Err.Raise vbObjectError + 2405, , "モック途中データが不正・存在しない、または入力が変更されています。"
End Sub

Public Function ExcelMockDummy(ByVal stage As Long) As String
    Dim text As String, body As String
    If stage < 1 Or stage > 6 Then Err.Raise vbObjectError + 2406, , "モック評価回が不正です。"
    If stage <= 5 Then
        text = EXCEL_MOCK_MARKER & vbLf & "審査員" & CStr(stage) & "のダミー批評です。構成の読みやすさと人物描写の改善点を示す固定文です。小説への実際の評価ではありません。"
    Else
        text = EXCEL_MOCK_MARKER & vbLf & "5人のダミー批評をまとめる総評のテスト文です。実際の審査結果ではありません。"
    End If
    body = "{""status"":""completed"",""output"":[{""type"":""message"",""role"":""assistant"",""content"":[{""type"":""output_text"",""text"":" & JsonQuote(text) & "}]}]}"
    ExcelMockDummy = ParseCritiqueResponse(body)
End Function

Public Function ExcelMockReview(ByVal stage As Long, ByVal instructions As String, ByVal inputText As String, ByVal stopAfter As Long, ByVal failAt As Long) As String
    If Len(checkpoint) = 0 Then Err.Raise vbObjectError + 2406, , "モックが初期化されていません。"
    If stage < 1 Or stage > 6 Then Err.Raise vbObjectError + 2406, , "モック評価回が不正です。"
    If stage <= completed Then
        ExcelMockReview = saved(stage)
        Exit Function
    End If
    If stage <> completed + 1 Then Err.Raise vbObjectError + 2406, , "モック評価順が不正です。"
    If stage = failAt Then Err.Raise EXCEL_MOCK_FAILURE, , "[モック] " & CStr(stage) & "回目のテスト用エラーです。"
    saved(stage) = ExcelMockDummy(stage)
    completed = stage
    SaveCheckpoint
    If stage = stopAfter Then Err.Raise EXCEL_MOCK_STOP, , "[モック] " & CStr(stage) & "/6 完了後に停止しました。再開時は同じ入力と保存先を選択してください。"
    ExcelMockReview = saved(stage)
End Function

Private Sub SaveCheckpoint()
    Dim text As String, i As Long, stream As Object, fso As Object, temporary As String
    Dim errorNumber As Long, errorText As String, created As Boolean
    text = "{""mode"":""singula-excel-mock-v1"",""fingerprint"":" & JsonQuote(fingerprint) & ",""results"":["
    For i = 1 To completed
        If i > 1 Then text = text & ","
        text = text & JsonQuote(saved(i))
    Next
    text = text & "]}"
    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    temporary = fso.BuildPath(fso.GetParentFolderName(checkpoint), fso.GetTempName)
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.SaveToFile temporary, 1
    created = True
    stream.Close
    If MoveFileExW(StrPtr(temporary), StrPtr(checkpoint), 9) = 0 Then Err.Raise vbObjectError + 2407, , "モック途中データを保存できません。保存先の権限を確認してください。"
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    If created Then fso.DeleteFile temporary
    On Error GoTo 0
    Err.Raise errorNumber, , errorText
End Sub
