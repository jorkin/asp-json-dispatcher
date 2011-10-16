<%
Class JsonEncoder
  Private buffer(), n, l, encoderMap
  Public RecordsetColumnInfo
  Sub Class_Initialize()
    n = 0
    l = 1024
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
      l = l + 1024
      ReDim Preserve buffer(l)
    End If
    buffer(n) = value
  End Sub

  Public Sub Encode(value)
    Dim n, c, char, key, item, valueType, bFirst
    bFirst = True
    valueType = TypeName(value)
    If encoderMap.Exists(valueType) Then
      Call encoderMap(valueType)(Me, value)
      Exit Sub
    End If
    Select Case valueType
'basic types
      Case "JsonEncoder"
        Write value.Result
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

'adodb.recordset
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
              If value.Fields(n).Name = "" Then
                Encode "Expr" & i
              Else
                Encode value.Fields(n).Name
              End If
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
      Case "Field"
        Write "{""name"":"
        Encode value.Name
        Write ",""definedSize"":"
        Encode value.DefinedSize
        Write ",""precision"":"
        Encode value.Precision
        Write ",""numericScale"":"
        Encode value.NumericScale
        Write ",""type"":"
        Encode value.Type
        Write "}"
      Case "Fields"
        c = value.Count
        Write "["
        For n = 0 To c - 1
          If n > 0 Then Write ","
          Encode value(n)
        Next
        Write "]"

'filesystem objects
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
        Write ",""type"":"
        Encode value.DriveType
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
        Write ",""shareName"":"
        Encode value.ShareName
        Write "}"

'string or type not presented
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
          ElseIf c = vbFormFeed Then
            Write "\f"
          ElseIf c = "/" Then
            Write "\/"
          Else
            char = AscW(c) And &HFFFF
            If char = &H0008 Then
              Write "\b"
            ElseIf char > &H00FF Or char < &H0020 Then
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

  Public Sub Clear()
    n = 0
    l = 1024
    ReDim buffer(l)
  End Sub

  Public Property Get Result()
    Result = Join(buffer, "")
  End Property
End Class
%>