throwError
==========
`throwError` &mdash; 抛出异常错误

说明
----
>     void throwError(ByVal throwErrorNumber, ByVal message)
> 抛出异常错误。

参数
----
> `throwErrorNumber`
>> **类型：**`integer`  
>> **说明：**异常错误编号。
>
> `message`
>> **类型：**`string`  
>> **说明：**异常错误信息。

返回值
------
> 没有返回值。

范例
----
>
    <%
    Dim message
    message = "系统内部错误。"
    Call SE.module("Error").throwError(2, message)
    %>