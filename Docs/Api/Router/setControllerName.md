setControllerName
=================
`setControllerName` &mdash; 设置控制器名称

说明
----
>     void setControllerName(string theControllerName)
> 设置控制器名称。

参数
----
> `theControllerName`
>> **类型：**`string`  
>> **说明：**控制器名称。

返回值
------
> 没有返回值。

范例
----
>
    <%
    SE.module("Router").setControllerName("Index")
    %>