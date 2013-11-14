setRedirectURL
==============
`setRedirectURL` &mdash; 设置重定向URL

说明
----
>     string setRedirectURL(ByVal urlString)
> 设置重定向URL。

参数
----
> `urlString`
>> **类型：**`string`  
>> **说明：**URL字符串。

返回值
------
> 没有返回值。

范例
----
>
    <%
    SE.module("Error").setRedirectURL("./")
    %>