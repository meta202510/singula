Attribute VB_Name = "Json"
Option Explicit

' Minimal recursive JSON reader. No external references required.
Private source As String
Private position As Long
Private depth As Long

Public Function JsonParse(ByVal text As String) As Object
    source = text
    position = 1
    depth = 0
    SkipSpace
    If Mid$(source, position, 1) <> "{" Then Fail
    Set JsonParse = ReadObject()
    SkipSpace
    If position <= Len(source) Then Fail
    source = ""
End Function

Public Function JsonQuote(ByVal text As String) As String
    Dim i As Long, c As String, n As Long, result As String
    result = Chr$(34)
    For i = 1 To Len(text)
        c = Mid$(text, i, 1)
        n = AscW(c)
        Select Case n
            Case 34: result = result & Chr$(92) & Chr$(34)
            Case 92: result = result & Chr$(92) & Chr$(92)
            Case 0 To 31: result = result & Chr$(92) & "u" & Right$("0000" & Hex$(n), 4)
            Case Else: result = result & c
        End Select
    Next
    JsonQuote = result & Chr$(34)
End Function

Private Sub SkipSpace()
    Do While position <= Len(source)
        Select Case Mid$(source, position, 1)
            Case " ", vbTab, vbCr, vbLf: position = position + 1
            Case Else: Exit Do
        End Select
    Loop
End Sub

Private Sub Expect(ByVal token As String)
    SkipSpace
    If Mid$(source, position, Len(token)) <> token Then Fail
    position = position + Len(token)
End Sub

Private Sub Fail()
    Err.Raise vbObjectError + 2100, "Json", "Invalid JSON at character " & CStr(position)
End Sub

Private Sub EnterContainer()
    depth = depth + 1
    If depth > 100 Then Fail
End Sub

Private Function ReadObject() As Object
    Dim result As Object, key As String
    EnterContainer
    Set result = CreateObject("Scripting.Dictionary")
    Expect "{"
    SkipSpace
    If Mid$(source, position, 1) <> "}" Then
        Do
            SkipSpace
            key = ReadString()
            Expect ":"
            If result.Exists(key) Then Fail
            result.Add key, ReadValue()
            SkipSpace
            If Mid$(source, position, 1) <> "," Then Exit Do
            position = position + 1
        Loop
    End If
    Expect "}"
    depth = depth - 1
    Set ReadObject = result
End Function

Private Function ReadArray() As Collection
    Dim result As New Collection
    EnterContainer
    Expect "["
    SkipSpace
    If Mid$(source, position, 1) <> "]" Then
        Do
            result.Add ReadValue()
            SkipSpace
            If Mid$(source, position, 1) <> "," Then Exit Do
            position = position + 1
        Loop
    End If
    Expect "]"
    depth = depth - 1
    Set ReadArray = result
End Function

Private Function ReadValue() As Variant
    SkipSpace
    Select Case Mid$(source, position, 1)
        Case "{": Set ReadValue = ReadObject()
        Case "[": Set ReadValue = ReadArray()
        Case Chr$(34): ReadValue = ReadString()
        Case "t": Expect "true": ReadValue = True
        Case "f": Expect "false": ReadValue = False
        Case "n": Expect "null": ReadValue = Null
        Case Else: ReadValue = ReadNumber()
    End Select
End Function

Private Function ReadString() As String
    Dim c As String, result As String, hexCode As String, n As Long, i As Long
    Expect Chr$(34)
    Do While position <= Len(source)
        c = Mid$(source, position, 1)
        position = position + 1
        If c = Chr$(34) Then
            ReadString = result
            Exit Function
        End If
        If c = Chr$(92) Then
            c = Mid$(source, position, 1)
            position = position + 1
            Select Case c
                Case Chr$(34), Chr$(92), "/": result = result & c
                Case "b": result = result & Chr$(8)
                Case "f": result = result & Chr$(12)
                Case "n": result = result & vbLf
                Case "r": result = result & vbCr
                Case "t": result = result & vbTab
                Case "u"
                    hexCode = Mid$(source, position, 4)
                    If Len(hexCode) <> 4 Then Fail
                    For i = 1 To 4
                        If InStr(1, "0123456789abcdef", Mid$(hexCode, i, 1), vbTextCompare) = 0 Then Fail
                    Next
                    n = CLng("&H" & hexCode)
                    If n > 32767 Then n = n - 65536
                    result = result & ChrW$(n)
                    position = position + 4
                Case Else: Fail
            End Select
        Else
            If AscW(c) >= 0 And AscW(c) < 32 Then Fail
            result = result & c
        End If
    Loop
    Fail
End Function

Private Function IsDigit() As Boolean
    Dim c As String
    c = Mid$(source, position, 1)
    IsDigit = (Len(c) = 1 And InStr(1, "0123456789", c, vbBinaryCompare) > 0)
End Function

Private Function ReadNumber() As Double
    Dim start As Long, c As String
    start = position
    If Mid$(source, position, 1) = "-" Then position = position + 1
    If Not IsDigit() Then Fail
    If Mid$(source, position, 1) = "0" Then
        position = position + 1
    Else
        Do While IsDigit(): position = position + 1: Loop
    End If
    If Mid$(source, position, 1) = "." Then
        position = position + 1
        If Not IsDigit() Then Fail
        Do While IsDigit(): position = position + 1: Loop
    End If
    c = Mid$(source, position, 1)
    If c = "e" Or c = "E" Then
        position = position + 1
        c = Mid$(source, position, 1)
        If c = "+" Or c = "-" Then position = position + 1
        If Not IsDigit() Then Fail
        Do While IsDigit(): position = position + 1: Loop
    End If
    ReadNumber = Val(Mid$(source, start, position - start))
End Function
