getAppName
==========
`getAppName` &mdash; 获取应用名称

说明
----
>     string getAppName(void)
> 获取应用名称。

参数
----
> 没有参数。

返回值
------
> **类型：**`string`  
> **说明：**应用名称。

范例
----
>
    <%
    Dim appName
    appName = SE.module("Router").getAppName
    %>