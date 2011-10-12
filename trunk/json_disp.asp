<%
Class JsonEncoder
  Private buffer(), n, l, encoderMap
  Public RecordsetColumnInfo
  Sub Class_Initialize()
    n = 0
    l = 100
    ReDim buffer(l)
    RecordsetColumnInfo = False
    Set encoderMap = Server.CreateObject("scripting.dictionary")
  End Sub

  Sub Class_Terminate()
    Set encoderMap = Nothing
  End Sub

  Public Sub AddEncoder(valueTypeName, encoder)
    encoderMap.Add valueTypeName, encoder
  End Sub

  Public Sub Write(value)
    n = n + 1
    If n > l Then
      l = l + 100
      ReDim Preserve buffer(l)
    End If
    buffer(n) = value
  End Sub

  Public Sub Encode(value)
    Dim n, c, char, key, item, valueType, bFirst : bFirst = True
    valueType = TypeName(value)
    If encoderMap.Exists(valueType) Then
      Call encoderMap(valueType)(Me, value)
      Exit Sub
    End If
    Select Case valueType
      Case "Boolean" : Write LCase(value)
      Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal" : Write value
      'Case "Date" : Write """Date(" : Write DateDiff("s", #1970-01-01#, value) * 1000 : Write ")"""
      Case "Empty", "Null", "Nothing" : Write "null"
      Case "Variant()"
        Write "["
        For Each item In value
          If bFirst Then
            bFirst = False
          Else
            Write ","
          End If
          Encode item
        Next
        Write "]"
      Case "Dictionary"
        Write "{"
        For Each key In value.Keys
          If bFirst Then
            bFirst = False
          Else
            Write ","
          End If
          Encode key
          Write ":"
          Encode value(key)
        Next
        Write "}"
      Case "Recordset"
        c = value.Fields.Count
        Write "["
        Do While Not value.Eof
          If bFirst Then
            bFirst = False
          Else
            Write ","
          End If
          If RecordsetColumnInfo Then
            Write "{"
            For n = 0 To c - 1
              If n > 0 Then Write ","
              Encode value.Fields(n).Name
              Write ":"
              Encode value(n).Value
            Next
            Write "}"
          Else
            Write "["
            For n = 0 To c - 1
              If n > 0 Then Write ","
              Encode value(n).Value
            Next
            Write "]"
          End If
        value.MoveNext : Loop
        Write "]"
      Case "File"
        Write "{""name"":"
        Encode value.Name
        Write ",""path"":"
        Encode value.Path
        Write ",""type"":"
        Encode value.Type
        Write ",""size"":"
        Encode value.Size
        Write ",""dateCreated"":"
        Encode value.DateCreated
        Write ",""dateLastModified"":"
        Encode value.DateLastModified
        Write ",""dateLastAccessed"":"
        Encode value.DateLastAccessed
        Write ",""attributes"":"
        Encode value.Attributes
        Write "}"
      Case "Files", "Folders"
        Write "["
        For Each item In value
          If bFirst Then
            bFirst = False
          Else
            Write ","
          End If
          Encode item
        Next
        Write "]"
      Case "Folder"
        Write "{""name"":"
        Encode value.Name
        Write ",""path"":"
        Encode value.Path
        Write ",""type"":"
        Encode value.Type
        Write ",""dateCreated"":"
        Encode value.DateCreated
        Write ",""dateLastModified"":"
        Encode value.DateLastModified
        Write ",""dateLastAccessed"":"
        Encode value.DateLastAccessed
        Write ",""attributes"":"
        Encode value.Attributes
        Write "}"
      Case "Drive"
        Write "{""name"":"
        Encode value.DriveLetter & ":"
        Write ",""path"":"
        Encode value.Path
        Write ",""totalSize"":"
        Encode value.TotalSize
        Write ",""availableSpace"":"
        Encode value.AvailableSpace
        Write ",""freeSpace"":"
        Encode value.FreeSpace
        Write ",""volumeName"":"
        Encode value.VolumeName
        Write ",""isReady"":"
        Encode value.IsReady
        Write ",""driveType"":"
        Encode value.DriveType
        Write ",""ShareName"":"
        Encode value.ShareName
        Write "}"
      Case "JsonEncoder"
        Write value.Json()
      Case Else
        Write """"
        For n = 1 To Len(value)
          c = Mid(value, n, 1)
          If c = """" Then
            Write "\"""
          ElseIf c = "\" Then
            Write "\\"
          ElseIf c = vbCr Then
            Write "\r"
          ElseIf c = vbLf Then
            Write "\n"
          ElseIf c = vbTab Then
            Write "\t"
          Else
            char = AscW(c) And &HFFFF
            If char > &H00FF Or char < &H0020 Then
              Write "\u"
              Write Right("0000" & Hex(char), 4)
            Else
              Write c
            End If
          End If
        Next
        Write """"
    End Select
  End Sub

  Public Sub Add(key, value)
    If n > 0 Then Write ","
    Encode Key
    Write ":"
    Encode value
  End Sub

  Public Sub Clear()
    n = 0
    l = 100
    ReDim buffer(l)
  End Sub

  Public Function Json()
    Json = "{" & Join(buffer, "") & "}"
  End Function
End Class

Class JsonDispatcher
  Dim acceptType, jsonCodePage, currentCodePage, jsonReturn
  Sub Class_Initialize()
    Set jsonReturn = new JsonEncoder
    jsonCodePage = 65001
    currentCodePage = Session.CodePage
  End Sub

  Sub Class_Terminate()
    Set jsonReturn = Nothing
  End Sub

  Public Sub AddEncoder(valueTypeName, encoder)
    jsonReturn.AddEncoder valueTypeName, encoder
  End Sub

  Public Sub AcceptParam(methodName, paramName)
    Session.CodePage = jsonCodePage
    If Request(paramName).Count = 0 Then
      Session.CodePage = currentCodePage
      Exit Sub
    End If
    Session.CodePage = currentCodePage
    Call Accept(methodName)
  End Sub

  Public Sub AcceptParamValue(methodName, paramName, paramValue)
    Dim n, bFound : bFound = False
    Session.CodePage = jsonCodePage
    If Request(paramName).Count = 0 Then
      Session.CodePage = currentCodePage
      Exit Sub
    End If
    For n = 1 To Request(paramName).Count
      If Request(paramName)(n) = paramValue Then
        bFound = True
        Exit For
      End If
    Next
    Session.CodePage = currentCodePage
    If bFound Then
      Call Accept(methodName)
    End If
  End Sub

  Public Sub AcceptForm(methodName, paramName)
    Session.CodePage = jsonCodePage
    If Request.Form(paramName).Count = 0 Then
      Session.CodePage = currentCodePage
      Exit Sub
    End If
    Session.CodePage = currentCodePage
    Call Accept(methodName)
  End Sub

  Public Sub AcceptFormValue(methodName, paramName, paramValue)
    Dim n, bFound : bFound = False
    Session.CodePage = jsonCodePage
    If Request.Form(paramName).Count = 0 Then
      Session.CodePage = currentCodePage
      Exit Sub
    End If
    For n = 1 To Request.Form(paramName).Count
      If Request.Form(paramName)(n) = paramValue Then
        bFound = True
        Exit For
      End If
    Next
    Session.CodePage = currentCodePage
    If bFound Then
      Call Accept(methodName)
    End If
  End Sub

  Public Sub Accept(methodName)
    On Error Resume Next
    Dim retVal, errNum, errSource, errDesc
    Session.CodePage = jsonCodePage
    retVal = GetRef(methodName)(jsonReturn)
    Session.CodePage = currentCodePage
    If Err.number <> 0 Then
      errNum = Err.number
      errSource = Err.Source
      errDesc = Err.Description
      Err.Clear
      On Error Goto 0
      Err.Raise errNum, errSource, errDesc & " - Err Raised While JsonDispatcher.Accept(" & methodName & ")"
    End If
    If Not retVal Then
      Exit Sub
    End If

    Response.ContentType = "application/json"
    Response.Write jsonReturn.Json()
    Response.End
  End Sub
End Class
%>