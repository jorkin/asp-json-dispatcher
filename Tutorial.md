## JsonEncoder ##
  * **Write(value)**
> > 결과값 반환을 위한 내부 버퍼에 문자열을 추가

  * **Encode(value)**
> > TypeName(value) 에 따라 JSON 문자열로 변환 후 Write 을 통해 내부 버퍼에 추가

  * **Add(key, value)**
> > 요청 처리후 반환될 JSON 결과에 key, value Pair를 추가

  * **Clear()**
> > 내부 버퍼를 비웁니다

  * **Json()**
> > 내부버퍼의 문자열을 조합하여, JSON 형식으로 변환된 결과값을 반환합니다

```
  Set encoder = new JsonEncoder
  encoder.Add "key1", "value1"
  encoder.Add "key2", 0.2
  encoder.Add "key3", Array(1, True, Null)
  Response.Write encoder.Json()

  => {"key1":"value1","key2":0.2,"key3":[1,true,null]}
```

## JsonDispatcher ##
  * **Accept(methodName)**
> > Function [[methodName](methodName.md)] 을 실행한 반환값이 False 가 아니면, JSON 결과값을 클라이언트에 전송하고 페이지실행을 중지(_Response.End_)합니다

  * **AcceptParam(methodName, paramName)**
> > Request(paramName) 이 존재한다면, _Accept(methodName)_ 을 실행합니다

  * **AcceptParamValue(methodName, paramName, paramValue)**
> > Request(paramName) 이 paramValue 와 같다면, _Accept(methodName)_ 을 실행합니다