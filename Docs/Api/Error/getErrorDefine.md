getErrorDefine
==============
`getErrorDefine` &mdash; 获取异常错误编号定义

说明
----
>     string getErrorDefine(ByVal errorNumber)
> 获取异常错误编号定义。

参数
----
> `errorNumber`
>> **类型：**`integer`  
>> **说明：**异常错误编号。

返回值
------
> **类型：**`string`  
> **说明：**异常错误编号定义。

范例
----
>
    <%
    Dim errorDefine
    errorDefine = SE.module("Error").getErrorDefine(0)
    Response.Write(errorDefine)
    %>
> 以上内容输出：
>
    系统正常