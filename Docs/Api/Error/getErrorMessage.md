getErrorMessage
===============
`getErrorMessage` &mdash; 获取异常错误消息。

说明
----
>     string getErrorMessage(void)
> 获取异常错误消息。

参数
----
> 没有参数。

返回值
------
> **类型：**`string`  
> **说明：**异常错误消息。

范例
----
>
    <%
    Dim message
    message = SE.module("Error").getErrorMessage
    %>