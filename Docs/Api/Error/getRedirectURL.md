getRedirectURL
==============
`getRedirectURL` &mdash; 获取重定向URL

说明
----
>     string getRedirectURL(void)
> 获取重定向URL。

参数
----
> 没有参数。

返回值
------
> **类型：**`string`  
> **说明：**重定向URL。

范例
----
>
    <%
    Dim redirectURL
    redirectURL = SE.module("Error").getRedirectURL
    %>