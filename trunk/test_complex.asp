<% Option Explicit %>
<!--#include file="JsonEncoder.asp"-->
<!--#include file="JsonDispatcher.asp"-->
<script language="jscript" runat="server">
var ExtJsonEncoder = {
  init: function(jsonDispatcher){
    for(var key in this){
      if(key == 'init') continue;
      jsonDispatcher.AddEncoder(key, this[key]);
    }
  }
, "Date": function(encoder, value){
    encoder.Write('"Date(' + new Date(value).getTime() + ')"');
  }
, "DOMDocument": function(encoder, value){
    encoder.Encode(value.xml);
  }
, "JScriptTypeInfo": function(encoder, value){
    if(value.constructor == Date){
      encoder.Write('"Date(' + value.getTime() + ')"');
      return;
    }
    if(value.constructor == Number){
      encoder.Write(isNaN(value) ? 'null' : value);
      return;
    }
    if(value.constructor == Array){
      encoder.Write('[');
      for(var i = 0, l = value.length; i < l; i++){
        if(i > 0) encoder.Write(',');
        encoder.Encode(value[i]);
      }
      encoder.Write(']');
      return;
    }
    encoder.Write('{');
    var n = 0;
    for(var key in value){
      if(typeof(value[key]) == 'function') continue;
      if(n > 0){
        encoder.Write(',');
      } else {
        n++;
      }
      encoder.Encode(key);
      encoder.Write(':');
      encoder.Encode(value[key]);
    }
    encoder.Write('}');
  }
, "Variant()": function(encoder, value){
    //For (adodb.recordset).GetRows
    var vbArr = new VBArray(value), jsArr = vbArr.toArray(), dimensions = vbArr.dimensions();
    if(dimensions > 1){
      for(var n = 1; n < dimensions; n++){
        var length = vbArr.ubound(n) + 1, arr = [], temp = [];
        for(var i = 0, l = jsArr.length; i < l; i++){
          if(i >= length && i % length == 0) {
            arr.push(temp);
            temp = [];
          }
          temp.push(jsArr[i]);
        }
        arr.push(temp);
        jsArr = arr;
      }
    }
    encoder.Encode(jsArr);
  }
};

function TestJSObject(){
  var o = new Object;
  o.random = Math.random();
  o.date = new Date();
  o.number = new Number("NaN number");
  o.text = "some \"test\" text\r\n from js object";
  return o;
}
</script>
<%
Function ComplexTest(result)
  Dim db, rs, xml, dic, jsObject, arr
  Select Case Request("mode")
    Case "test1"
      result.Add "data", Now
      result.Add "type", "Date"
    Case "test2"
      Set xml = Server.CreateObject("msxml.domdocument")
      xml.LoadXML "<data><str><![CDATA[테스트 TEST]]></str></data>"
      result.Add "data", xml
      result.Add "type", "XML DOMDocument"
      Set xml = Nothing
    Case "test3"
      Set dic = Server.CreateObject("scripting.dictionary")
      dic.Add "first", "First String"
      dic.Add "second", 1
      result.Add "data", dic
      Set dic = Nothing
      result.Add "type", "Dictionary"
    Case "test4"
      result.Add "data", Array("A", 0.5, False)
      result.Add "nested", Array("B", 1, True, Array("C", 100))
      result.Add "type", "Array And Nested Array"
    Case "test5"
      Set jsObject = TestJSObject()
      result.Add "data", jsObject
      result.Add "type", "js object"
    Case "test6"
      Set db = Server.CreateObject("adodb.connection")
      db.Open "Provider=Microsoft.JET.OLEDB.4.0;Data Source=""" & Server.MapPath(".") & """;Extended Properties=""Text;HDR=YES;FMT=Delimited"""
      Set rs = db.Execute("SELECT Name, Email FROM [dummy_data.csv]")
      result.Encoder.RecordsetColumnInfo = True
      result.Add "columns", rs.Fields
      result.Add "data", rs
      result.Add "type", "Recordset"
      rs.Close
      db.Close
      Set rs = Nothing
      Set db = Nothing
    Case "test7"
      Set db = Server.CreateObject("adodb.connection")
      db.Open "Provider=Microsoft.JET.OLEDB.4.0;Data Source=""" & Server.MapPath(".") & """;Extended Properties=""Text;HDR=YES;FMT=Delimited"""
      Set rs = db.Execute("SELECT * FROM [dummy_data.csv]")
      If Not rs.Eof Then arr = rs.GetRows()
      rs.Close
      db.Close
      Set rs = Nothing
      Set db = Nothing
      result.Add "data", arr
      result.Add "type", "Multi-dimensional Array"
    Case Else
      ComplexTest = False
      Exit Function
  End Select
  ComplexTest = True
End Function

ExtJsonEncoder.init JsonDispatcher

Call JsonDispatcher.AcceptParam(GetRef("ComplexTest"), "mode")
%>
<!doctype html>
<html lang="ko">
<head>
  <title>ajax complex type test</title>
  <meta http-equiv="content-type" content="text/html; charset=utf-8">
  <script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.5.2/jquery.min.js"></script>
  <script type="text/javascript">
  $(function(){
    $('body').ajaxError(function(ev, xhr, textStatus, errorThrown) {
      $('<div></div>').text("error - " + xhr.responseText).appendTo('#response');
    }).ajaxSuccess(function(ev, xhr) {
      $('<div></div>').text("success - " + xhr.responseText).appendTo('#response');
    });
  });

  $(function(){
    $('#ajaxtest').click(function(){

      $('#result, #response').empty();

      $.post('test_complex.asp',{mode:'test1'},function(result){
        $('<div>' + eval('new ' + result.data).toLocaleString() + ' <em>' + result.type + '</em></div>').appendTo('#result');
      });

      $.post('test_complex.asp',{mode:'test2'},function(result){
        if($.browser.msie){
          var x = new ActiveXObject('microsoft.xmldom');
          x.loadXML(result.data);
          result.data = x;
        } else if(window.DOMParser) {
          result.data = new DOMParser().parseFromString(result.data,'text/xml');
        }
        $('<div>' + $(result.data).find('str').text() + ' <em>' + result.type + '</em></div>').appendTo('#result');
      });

      $.post('test_complex.asp',{mode:'test3'},function(result){
        $('<div>' + result.data.first + ',' + result.data.second + ' <em>' + result.type + '</em></div>').appendTo('#result');
      });

      $.post('test_complex.asp',{mode:'test4'},function(result){
        $('<div>data : [' + result.data.join(', ') + '], nested : [' + result.nested.data.join(', ') + '] <em>' + result.type + '</em></div>').appendTo('#result');
      });

      $.post('test_complex.asp',{mode:'test5'},function(result){
        $('<div>data.random : ' + result.data.random + ', data.date : ' + result.data.date + ', data.number : ' + result.data.number + ', data.text : ' + result.data.text + ' <em>' + result.type + '</em></div>').appendTo('#result');
      });

      $.post('test_complex.asp',{mode:'test6'},function(result){
        $('<div>' +
          $(result.data).map(function(){
            return '<span>' + this.Name + ' (' + this.Email + ')</span>';
          }).toArray().join(', ') + ' <em>' + result.type + '</em></div>'
        ).appendTo('#result');
      });

      $.post('test_complex.asp',{mode:'test7'},function(result){
        $('<div>' + result.data + ' <em>' + result.type + '</em></div>').appendTo('#result');
      });

    });
  });
  </script>
  <style type="text/css">
  body,input,button{font-family:Lucida console;font-size:10pt;line-height:1.5;}
  em{color:blue;}
  </style>
</head>
<body>
  <button type="button" id="ajaxtest">TEST</button>
  <hr>
  <div id="result"></div>
  <hr>
  <div id="response"></div>
</body>
</html>