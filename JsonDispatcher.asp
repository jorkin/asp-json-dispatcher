<%
'Requires JsonEncoder.asp

Class JsonResultClass
  Private enc, buffer(), n, l
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

  Sub Class_Terminate()
    Set jsonResult = Nothing
  End Sub

  Public Sub AddEncoder(valueTypeName, encoder)
    jsonResult.Encoder.AddEncoder valueTypeName, encoder
  End Sub

  Public Sub AcceptParam(method, paramName)
    Session.CodePage = jsonCodePage
    If Request(paramName).Count = 0 Then
      Session.CodePage = currentCodePage
      Exit Sub
    End If
    Session.CodePage = currentCodePage
    Call Accept(method)
  End Sub

  Public Sub AcceptParamValue(method, paramName, paramValue)
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
      Call Accept(method)
    End If
  End Sub

  Public Sub AcceptForm(method, paramName)
    Session.CodePage = jsonCodePage
    If Request.Form(paramName).Count = 0 Then
      Session.CodePage = currentCodePage
      Exit Sub
    End If
    Session.CodePage = currentCodePage
    Call Accept(method)
  End Sub

  Public Sub AcceptFormValue(method, paramName, paramValue)
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
      Call Accept(method)
    End If
  End Sub

  Public Sub Accept(method)
    Dim retVal, dummy
    If TypeName(method) = "String" Then Set method = GetRef(method)
    Session.CodePage = jsonCodePage
    dummy = Request("dummy for codepage conversion")
    Session.CodePage = currentCodePage
    retVal = method(jsonResult)
    If Not retVal Then
      Exit Sub
    End If

    Response.ContentType = "application/json"
    Response.Write jsonResult.Result
    Response.End
  End Sub
End Class

Dim JsonDispatcher
Set JsonDispatcher = new JsonDispatcherClass
%>