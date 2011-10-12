<!--#include file="json_disp.asp"-->
<%
Sub EncodeFileSystemObjectSimple(result, value)
  result.Write "{"
  result.Write """name"":"
  result.Encode value.Name
  result.Write ",""path"":"
  result.Encode value.Path
  result.Write "}"
End Sub

Function Test1(result)
  If Request("name") = "" Then
    result.Add "error", "Type Your Name"
  Else
    result.Add "error", Null
    result.Add "data", "Hello " & Request("name")
    result.Add "now", Now
  End If

  Test1 = True
End Function

Function Test2(result)
  Dim path
  path = Request.Form("path")
  If path = "" Then path = Server.MapPath("/")

  'add new external encoder
  result.AddEncoder "Folder", GetRef("EncodeFileSystemObjectSimple")

  Set fso = Server.CreateObject("scripting.filesystemobject")
  If fso.DriveExists(path) Then
    Set folder = fso.GetDrive(path).RootFolder
  Else
    Set folder = fso.GetFolder(path)
  End If
  result.Add "folder", folder
  result.Add "folders", folder.SubFolders
  result.Add "files", folder.Files
  Set folder = Nothing
  Set fso = Nothing

  Test2 = True
End Function

Function AjaxTest1(result)
  result.Add "ip", Request.ServerVariables("REMOTE_ADDR")
  result.Add "data", Array(1, "문자열", NOW)

  AjaxTest1 = True
End Function

Set json = new JsonDispatcher
Call json.AcceptParam("Test1", "name")
Call json.AcceptForm("Test2", "path")
Call json.AcceptParamValue("AjaxTest1", "mode", "test1")
Set json = Nothing
%>
<!doctype html>
<html lang="ko">
<head>
  <title>ajax test</title>
  <meta http-equiv="content-type" content="text/html; charset=euc-kr">
  <script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.5.2/jquery.min.js"></script>
  <script type="text/javascript" src="http://ajax.aspnetcdn.com/ajax/jquery.templates/beta1/jquery.tmpl.min.js"></script>
  <script type="text/javascript">
  $(function(){
    $('body').ajaxError(function(ev, xhr, textStatus, errorThrown) {
      $('#response').text("error - " + xhr.responseText);
    }).ajaxSuccess(function(ev, xhr) {
      $('#response').text("success - " + xhr.responseText);
    });
  });

  $(function(){
    $('#ajaxtest').click(function(){
      $.post('test_euckr.asp',{mode:'test1'},function(result){
        alert(result.ip);
        alert(result.data.join('\n'));
      },'json');
    });

    $('#test1').click(function(){
      $.post('test_euckr.asp',{name:$('#name').val()},function(result){
        if(result.error)
          alert(result.error);
        else
          $('#result').html(result.data + '<br>' + result.now);
      },'json');
    });

    $('div.folder').live('click',function(){
      getList($.attr(this,'title'));
    });
    $('#test2').click(function(){
      getList('')
    });
    var getList = function(path){
      $.post('test_euckr.asp',{path:path},function(result){
        $('#result').html(
          '<strong>' + result.folder.path + '</strong>' +
          '<div class="folder" title="' + result.folder.path.split('\\').slice(0, -1).join('\\') + '\\">..</div>' +
          $(result.folders).map(function(){
            return '<div class="folder" title="' + this.path + '">' + this.name + '</div>';
          }).toArray().join('') +
          $(result.files).map(function(){
            return '<div class="file">' + this.name + '<em>' + (this.size / 1000).toFixed(2) + 'Kb ' + this.type + '</em></div>';
          }).toArray().join('')
        );
      },'json');
    }
  });
  </script>
  <style type="text/css">
  body,input,button{font-size:9pt;line-height:1.5;}
  .folder{cursor:pointer;}
  .folder:hover{background:blue;color:#fff;}
  .file em{margin-left:20px;color:#777;}
  xmp{font-size:10pt;font-family:Lucida Console,consolas;}
  dt{font-weight:bold;}
  </style>
</head>
<body>
  <h1>예제</h1>
  <fieldset>
    <legend>server.asp</legend>
    <xmp>
    Function AjaxTest1(result)
      result.Add "ip", Request.ServerVariables("REMOTE_ADDR")
      result.Add "data", Array(1, "문자열", NOW)

      AjaxTest1 = True
    End Function
    
    Set json = new JsonDispatcher
    Call json.AcceptParamValue("AjaxTest1", "mode", "test1")
    </xmp>
  </fieldset>
  <fieldset>
    <legend>client-jquery.htm</legend>
    <xmp>
    <script type="text/javascript">
    $(function(){
      $('#ajaxtest').click(function(){
        $.post('server.asp',{mode:'test1'},function(result){
          alert(result.ip);
          alert(result.data.join('\n'));
        },'json');
      });
    });
    </script>
    <button type="button" id="ajaxtest">test1</button>
    </xmp>
  </fieldset>
  <h1>데모</h1>
  <button type="button" id="ajaxtest">예제 실행</button><br>
  <input type="text" id="name" value=""><button type="button" id="test1">Hello 테스트</button><br>
  <button type="button" id="test2">파일시스템 테스트</button>
  <hr>
  <div id="result"></div>
  <hr>
  <div id="response"></div>
</body>
</html>