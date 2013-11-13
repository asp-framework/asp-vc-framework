getDBConnectionStatus
=====================
`getDBConnectionStatus` &mdash; 获取数据库连接状态

说明
----
>     integer getDBConnectionStatus(void)
> 获取数据库连接状态。

参数
----
> 没有参数。

返回值
------
> **类型：**`integer`  
> **说明：**数据库连接状态。(默认值为`0`。)
>> 关闭：`0`；  
>> 开启：`1`。

范例
----
>
    <%
    Dim connectionStatus
    connectionStatus = SE.module("DB").getDBConnectionStatus
    %>