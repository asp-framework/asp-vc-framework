setDatabaseInfo
===============
`setDatabaseInfo` &mdash; 设置数据库信息

说明
----
>     void setDatabaseInfo(object dbConnection)
> 设置数据库信息。(设置的信息将会在 **数据库信息** 面版中显示。)

参数
----
> `dbConnection`
>> **类型：** `object`  
>> **说明：**数据库连接。

返回值
------
> 没有返回值。

范例
----
>
    <%
    SE.module("DB").open()
    SE.module("Debugging").setDatabaseInfo( _
        SE.module("DB").getDBConnection _
    )
    SE.module("DB").close()
    %>