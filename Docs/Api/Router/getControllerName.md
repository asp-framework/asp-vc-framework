getControllerName
=================
`getControllerName` &mdash; 获取控制器名称

说明
----
>     string getControllerName(void)
> 获取控制器名称。

参数
----
> 没有参数。

返回值
------
> **类型：**`string`  
> **说明：**控制器名称。

范例
----
>
    <%
    Dim controllerName
    controllerName = SE.module("Router").getControllerName
    %>