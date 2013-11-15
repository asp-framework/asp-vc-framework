getActionName
=============
`getActionName` &mdash; 获取动作名称

说明
----
>     string getActionName(void)
> 获取动作名称。

参数
----
> 没有参数。

返回值
------
> **类型：**`string`  
> **说明：**动作名称。

范例
----
>
    <%
    Dim actionName
    actionName = SE.module("Router").getActionName
    %>