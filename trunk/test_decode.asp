<% Option Explicit %>
<!--#include file="JsonEncoder.asp"-->
<!--#include file="JsonDecoder.asp"-->
<!--#include file="JsonDispatcher.asp"-->
<%
Function DecodeTest(result)
  On Error Resume Next
  Dim decoder, encoder, jsonObject, isDecoded
  Set decoder = new JsonDecoder
  isDecoded = decoder.Decode(Request.Form("json"), jsonObject)
  Set decoder = Nothing

  If Err.number <> 0 Then
    result.Add "error", Err.Description
    Err.Clear
  Else
    Set encoder = new JsonEncoder
    encoder.Encode jsonObject  
    result.Add "error", Null
    result.Add "data", encoder.Result
    Set encoder = Nothing
  End If

  DecodeTest = True
End Function

Call JsonDispatcher.AcceptForm("DecodeTest", "json")
%>
<!doctype html>
<html lang="en">
<head>
  <title>JSON decode test</title>
  <meta http-equiv="content-type" content="text/html; charset=utf-8">
  <script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.5.2/jquery.min.js"></script>
  <script type="text/javascript" src="http://ajax.aspnetcdn.com/ajax/jquery.templates/beta1/jquery.tmpl.min.js"></script>
  <script type="text/javascript">
  $(function(){
    $('body').ajaxError(function(ev, xhr, textStatus, errorThrown) {
      $('#response').text("error - " + xhr.responseText);
    });
  });

  $(function(){
    $('#decode').click(function(){
      $.post('test_decode.asp',{json:$('#json').val()},function(result){
        if(result.error) {
          $('#result').text('Invalid JSON Expression - ' + result.error);
          $('#response').empty();
        } else {
          $('#result').text('Valid JSON Expression');
          $('#response').text(result.data);
        }
      },'json');
    });
  });
  </script>
  <style type="text/css">
  body,input,textarea,button{font-size:10pt;line-height:1.5;font-family:Consolas, Lucida Console;}
  </style>
</head>
<body>
  <fieldset>
    <legend>JSON Decode Test</legend>
    <textarea id="json" rows="10" cols="100">{"hello":"world"}</textarea><br>
    <button type="decode" id="decode">Decode</button>
  </fieldset>
  <hr>
  <strong>Validate JSON Result</strong>
  <div id="result"></div>
  <hr>
  <strong>Result Of JsonEncoder.Encode(JsonDecoder.Decode({your input}))</strong>
  <div id="response"></div>
</body>
</html>