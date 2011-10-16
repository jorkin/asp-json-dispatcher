<%
Class JsonDecoder
  Private buffer(), bufferPos, bufferLen, text, textPos, textLen, trimReg
  Sub Class_Initialize()
    Set trimReg = new RegExp
    trimReg.Pattern = "^\s+|\s+$"
    trimReg.Global = True
    text = ""
    textLen = 0
    textPos = 0
    Call ClearBuffer()
  End Sub

  Sub Class_Terminate()
    Set trimReg = Nothing
    Call ClearBuffer()
  End Sub

  Private Function Trim2(value)
    Trim2 = trimReg.Replace(value, "")
  End Function

  Private Function ReadString()
    Dim c
    Do While textPos < textLen
      c = ReadChar()
      If c = "\" Then
        c = ReadChar()
        Select Case c
          Case """"
            Call AddBuffer("""")
          Case "t"
            Call AddBuffer(vbTab)
          Case "r"
            Call AddBuffer(vbCr)
          Case "n"
            Call AddBuffer(vbLf)
          Case "f"
            Call AddBuffer(vbFormFeed)
          Case "b"
            Call AddBuffer(Chr(8))
          Case "u"
            c = ReadChars(5)
            If Len(c) <> 5 Then
              Err.Raise 1002, "JsonDecoder", Left(text, textPos) & " <= Expecting \u[0-9a-fA-F]{4}"
            End If
            Call AddBuffer(ChrW(CLng("&H" & Right(c, 4))))
          Case Else
            Call AddBuffer("\" & c)
        End Select
      ElseIf c = """" Then
        Exit Do
      Else
        Call AddBuffer(c)
      End If
    Loop
    If c <> """" Then Err.Raise 1002, "JsonDecoder", Left(text, textPos) & " <= Expecting '""'"
    ReadString = GetBuffer()
    Call ClearBuffer()
  End Function

  Private Function ReadNumber()
    Dim c, isDoubleType
    isDoubleType = False
    textPos = textPos - 1
    Do While textPos < textLen
      c = ReadChar()
      Select Case c
        Case "e", "E", "."
          Call AddBuffer(c)
          isDoubleType = True
        Case "-", "+", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
          Call AddBuffer(c)
        Case Else
          textPos = textPos - 1
          Exit Do
      End Select
    Loop
    If isDoubleType Then
      ReadNumber = CDbl(GetBuffer())
    Else
      ReadNumber = CLng(GetBuffer())
    End If
    Call ClearBuffer()
  End Function

  Private Function ReadBoolean()
    Call ClearBuffer()
    Dim c
    textPos = textPos - 1
    c = ReadChar()
    If c = "t" Then
      c = ReadChars(4)
    ElseIf c = "f" Then
      c = ReadChars(5)
    Else
      Err.Raise 1002, "JsonDecoder", Left(text, textPos) & " <= Expecting true or false"
    End If

    If c = "true" Then
      ReadBoolean = True
    ElseIf c = "false" Then
      ReadBoolean = False
    Else
      Err.Raise 1002, "JsonDecoder", Left(text, textPos) & " <= Expecting true or false"
    End If
  End Function

  Private Function ReadNull()
    Call ClearBuffer()
    Dim c
    c = ReadChars(4)
    If c = "null" Then
      ReadNull = Null
    Else
      Err.Raise 1002, "JsonDecoder", Left(text, textPos) & " <= Expecting null"
    End If
  End Function

  Private Function ReadBlock(blockStart, blockEnd)
    Dim c, depth, inString
    depth = 0
    inString = False
    Do While textPos < textLen
      c = ReadChar()
      If c = """" Then
        If Not inString Then
          inString = True
        ElseIf inString And GetLastBuffer() <> "\" Then
          inString = False
        End If
      ElseIf c = blockStart Then
        If Not inString Then
          depth = depth + 1
        End If
      ElseIf c = blockEnd Then
        If Not inString Then
          If depth = 0 Then
            Exit Do
          Else
            depth = depth - 1
          End If
        End If
      End If
      Call AddBuffer(c)
    Loop
    If c <> blockEnd Then Err.Raise 1002, "JsonDecoder", Left(text, textPos) & " <= Expecting '" & blockEnd & "'"
    ReadBlock = GetBuffer()
    Call ClearBuffer()
  End Function

  Private Function ReadChar()
    Dim c
    textPos = textPos + 1
    c = Mid(text, textPos, 1)
    ReadChar = c
    If textPos > textLen Then Err.Raise 1002, "JsonDecoder", "text EOF"
  End Function

  Private Function ReadChars(ByVal length)
    If textPos + length > textLen Then
      length = textLen - textPos + 1
    End If
    ReadChars = Mid(text, textPos, length)
    textPos = textPos + length - 1
  End Function

  Private Sub AddBuffer(c)
    buffer(bufferPos) = c
    bufferPos = bufferPos + 1
    If bufferPos > bufferLen Then
      bufferLen = bufferLen + 1024
      ReDim Preserve buffer(bufferLen)
    End If
  End Sub

  Private Function GetBuffer()
    GetBuffer = Join(buffer, "")
  End Function

  Private Sub ClearBuffer()
    bufferPos = 0
    bufferLen = 1024
    ReDim buffer(bufferLen)
  End Sub

  Private Function GetLastBuffer()
    If bufferPos = 0 Then Exit Function
    GetLastBuffer = buffer(bufferPos - 1)
  End Function

  Public Property Get IsDictionary()
    IsDictionary = isDictionaryResult
  End Property

  Public Function Decode(value, ByRef result)
    Decode = False
    Call ClearBuffer()
    text = Trim2(value)
    textLen = Len(text)
    textPos = 0
    Dim c, temp
    c = ReadChar()
    If c = "{" Then
      temp = ReadBlock("{", "}")
      If textPos < textLen Then Err.Raise 1002, "JsonDecoder", Left(text, textPos + 1) & " <= Expecting EOF"
      Set result = DecodeMap(temp)
      Decode = True
    ElseIf c = "[" Then
      temp = ReadBlock("[", "]")
      If textPos < textLen Then Err.Raise 1002, "JsonDecoder", Left(text, textPos + 1) & " <= Expecting EOF"
      result = DecodeArray(temp)
      Decode = True
    Else
      Err.Raise 1002, "JsonDecoder", Left(text, textPos) & " <= Expecting '{' or '['"
    End If
  End Function

  Public Function DecodeArray(value)
    Dim arr(), j_flag, J_NONE, J_VALUE
    J_NONE = 0
    J_VALUE = 2
    ReDim arr(1024)
    Call ClearBuffer()
    text = Trim2(value)
    textLen = Len(text)
    textPos = 0
    j_flag = J_NONE
    
    Dim c, arrLen, decoder, temp
    arrLen = -1
    Do While textPos < textLen
      c = ReadChar()
      If c = "," Then
        If j_flag = J_NONE Then Err.Raise 1002, "JsonDecoder", Left(text, textPos) & " <= Expecting [^,]"
        j_flag = J_NONE
      ElseIf c = " " Or c = vbTab Or c = vbCr Or c = vbLf Then
        'skip space char
      Else
        If j_flag = J_VALUE Then Err.Raise 1002, "JsonDecoder", Left(text, textPos) & " <= Expecting ','"
        Select Case c
          Case "{"
            temp = ReadBlock("{", "}")
            Set decoder = new JsonDecoder
            arrLen = arrLen + 1
            Set arr(arrLen) = decoder.DecodeMap(temp)
            Set decoder = Nothing
            j_flag = J_VALUE
          Case "["
            temp = ReadBlock("[", "]")
            Set decoder = new JsonDecoder
            arrLen = arrLen + 1
            arr(arrLen) = decoder.DecodeArray(temp)
            Set decoder = Nothing
            j_flag = J_VALUE
          Case """"
            arrLen = arrLen + 1
            arr(arrLen) = ReadString()
            j_flag = J_VALUE
          Case "e", "E", "+", "-", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
            arrLen = arrLen + 1
            arr(arrLen) = ReadNumber()
            j_flag = J_VALUE
          Case "t", "f"
            arrLen = arrLen + 1
            arr(arrLen) = ReadBoolean()
            j_flag = J_VALUE
          Case "n"
            arrLen = arrLen + 1
            arr(arrLen) = ReadNull()
            j_flag = J_VALUE
          Case Else
            Err.Raise 1002, "JsonDecoder", Left(text, textPos) & " <= Expecting [,\:\""\{\[]"
        End Select
      End If
    Loop
    If j_flag = J_NONE Then Err.Raise 1002, "JsonDecoder", text & " <= Expecting [^,]"
    ReDim Preserve arr(arrLen)
    DecodeArray = arr
  End Function

  Public Function DecodeMap(value)
    Dim map, j_flag, J_NONE, J_KEY, J_VALUE
    Set map = Server.CreateObject("scripting.dictionary")
    J_NONE = 0
    J_KEY = 1
    J_VALUE = 2
    j_flag = J_NONE
    Call ClearBuffer()
    text = Trim2(value)
    textLen = Len(text)
    textPos = 0

    Dim c, lastKey, lastValue, decoder, temp

    Do While textPos < textLen
      c = ReadChar()

      Select Case c
        Case "{"
          If j_flag = J_VALUE Then
            temp = ReadBlock("{", "}")
            Set decoder = new JsonDecoder
            map.Add lastKey, decoder.DecodeMap(temp)
            Set decoder = Nothing
          Else
            Err.Raise 1002, "JsonDecoder", Left(text, textPos) & " <= Expecting ':'"
          End If
        Case "["
          If j_flag = J_VALUE Then
            temp = ReadBlock("[", "]")
            Set decoder = new JsonDecoder
            map.Add lastKey, decoder.DecodeArray(temp)
            Set decoder = Nothing
          Else
            Err.Raise 1002, "JsonDecoder", Left(text, textPos) & " <= Expecting ':'"
          End If
        Case """"
          If j_flag = J_NONE Then
            lastKey = ReadString()
            j_flag = J_KEY
          ElseIf j_flag = J_VALUE Then
            lastValue = ReadString()
            map.Add lastKey, lastValue
          End If
        Case ":"
          If j_flag = J_KEY Then
            j_flag = J_VALUE
          Else
            Err.Raise 1002, "JsonDecoder", Left(text, textPos) & " <= Expecting [^:]"
          End If
        Case ","
          If j_flag = J_VALUE Then
            j_flag = J_NONE
          Else
            Err.Raise 1002, "JsonDecoder", Left(text, textPos) & " <= Expecting [^,]"
          End If
        Case "e", "E", "+", "-", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
          If j_flag = J_VALUE Then
            map.Add lastKey, ReadNumber()
          Else
            Err.Raise 1002, "JsonDecoder", Left(text, textPos) & " <= Expecting [^+\-eE0-9]"
          End If
        Case "t", "f"
          If j_flag = J_VALUE Then
            map.Add lastKey, ReadBoolean()
          Else
            Err.Raise 1002, "JsonDecoder", Left(text, textPos) & " <= Expecting '""'"
          End If
        Case "n"
          If j_flag = J_VALUE Then
            map.Add lastKey, ReadNull()
          Else
            Err.Raise 1002, "JsonDecoder", Left(text, textPos) & " <= Expecting '""'"
          End If
        Case " ", vbTab, vbCr, vbLf
          'skip space char
        Case Else
          Err.Raise 1002, "JsonDecoder", Left(text, textPos) & " <= Expecting [,\:\""\{\[]"
      End Select
    Loop
    If c = "," Then Err.Raise 1002, "JsonDecoder", text & " <= Expecting [^,]"
    Set DecodeMap = map
  End Function
End Class
%>