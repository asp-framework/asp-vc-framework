getError
========
`getError` &mdash; 获取当前异常错误编号

说明
----
>     integer getError(void)
> 获取当前异常错误编号。

参数
----
> 没有参数。

返回值
------
> **类型：**`integer`  
> **说明：**异常错误编号。(默认值为`0`。)

范例
----
>
    <%
    Dim errorNumber
    errorNumber = SE.module("Error").getError
    %>