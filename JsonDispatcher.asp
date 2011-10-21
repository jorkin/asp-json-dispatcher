<%
'Requires JsonEncoder.asp

Class JsonResultClass
  Private enc, buffer(), n, l
  
  Public Property Get Version()
    Version = "0.1.0"
  End Property
  
  Sub Class_Initialize()
    Call Clear()
    Set enc = new JsonEncoder
  End Sub

  Sub Class_Terminate()
    Set enc = Nothing
  End Sub

  Public Sub Clear()
    n = 0
    l = 256
    ReDim buffer(l)
  End Sub

  Private Sub AddBuffer(value)
    buffer(n) = value
    n = n + 1
    If n > l Then
      l = l + 256
      ReDim Preserve buffer(l)
    End If
  End Sub

  Public Sub Add(key, value)
    If n > 0 Then Call AddBuffer(",")
    Call enc.Clear()
    Call enc.Encode(key)
    Call AddBuffer(enc.Result)
    Call AddBuffer(":")
    Call enc.Clear()
    Call enc.Encode(value)
    Call AddBuffer(enc.Result)
    Call enc.Clear()
  End Sub

  Public Property Get Encoder()
    Set Encoder = enc
  End Property

  Public Property Get Result()
    Result = "{" & Join(buffer, "") & "}"
  End Property
End Class

Class JsonDispatcherClass
  Private acceptType, jsonCodePage, currentCodePage, jsonResult
  Sub Class_Initialize()
    Set jsonResult = new JsonResultClass
    jsonCodePage = 65001
    currentCodePage = Session.CodePage
  End Sub
  
  Public Property Let CodePage(value)
    jsonCodePage = value
  End Property
  
  Public Property Get CodePage()
    CodePage = jsonCodePage
  End Property
  
  Private Sub changeCodePage()
    If jsonCodePage <> currentCodePage Then Session.CodePage = jsonCodePage
  End Sub
  
  Private Sub restoreCodePage()
    If jsonCodePage <> currentCodePage Then Session.CodePage = currentCodePage
  End Sub

  Sub Class_Terminate()
    Set jsonResult = Nothing
  End Sub

  Public Sub AddEncoder(valueTypeName, encoder)
    jsonResult.Encoder.AddEncoder valueTypeName, encoder
  End Sub

  Public Sub AcceptParam(method, paramName)
    changeCodePage
    If Request(paramName).Count = 0 Then
      restoreCodePage
      Exit Sub
    End If
    restoreCodePage
    Call Accept(method)
  End Sub

  Public Sub AcceptParamValue(method, paramName, paramValue)
    Dim n, bFound : bFound = False
    changeCodePage
    If Request(paramName).Count = 0 Then
      restoreCodePage
      Exit Sub
    End If
    For n = 1 To Request(paramName).Count
      If Request(paramName)(n) = paramValue Then
        bFound = True
        Exit For
      End If
    Next
    restoreCodePage
    If bFound Then
      Call Accept(method)
    End If
  End Sub

  Public Sub AcceptForm(method, paramName)
    changeCodePage
    If Request.Form(paramName).Count = 0 Then
      restoreCodePage
      Exit Sub
    End If
    restoreCodePage
    Call Accept(method)
  End Sub

  Public Sub AcceptFormValue(method, paramName, paramValue)
    Dim n, bFound : bFound = False
    changeCodePage
    If Request.Form(paramName).Count = 0 Then
      restoreCodePage
      Exit Sub
    End If
    For n = 1 To Request.Form(paramName).Count
      If Request.Form(paramName)(n) = paramValue Then
        bFound = True
        Exit For
      End If
    Next
    restoreCodePage
    If bFound Then
      Call Accept(method)
    End If
  End Sub

  Public Sub Accept(method)
    Dim retVal, dummy
    If TypeName(method) = "String" Then Set method = GetRef(method)
    changeCodePage
    dummy = Request("dummy_for_codepage_conversion")
    restoreCodePage
    retVal = method(jsonResult)
    If Not retVal Then
      Exit Sub
    End If

    Response.Clear
    Response.ContentType = "application/json"
    Response.Write jsonResult.Result
    Response.End
  End Sub
End Class

Dim JsonDispatcher
Set JsonDispatcher = new JsonDispatcherClass
%>