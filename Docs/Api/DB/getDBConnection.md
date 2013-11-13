getDBConnection
===============
`getDBConnection` &mdash; 获取数据库连接

说明
----
>     object getDBConnection(void)
> 获取数据库连接。

参数
----
> 没有参数。

返回值
------
> **类型：**`object`  
> **说明：**数据库连接。

范例
----
>
    <%
    Dim connection
    connection = SE.module("DB").getDBConnection
    %>