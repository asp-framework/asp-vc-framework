executeSql
==========
`executeSql` &mdash; 执行SQL操作

说明
----
>     recordset executeSql(string sqlString)
> 执行SQL操作。

参数
----
> `sqlString`
>> **类型：**`string`  
>> **说明：**SQL语句。

返回值
------
> **类型：**`recordset`  
> **说明：**数据表记录集。

范例
----
>
    <!-- 配置文件 -->
    <?xml version="1.0" encoding="UTF-8"?>
    <SEConfigs>
        ...
        <DB>
            <type>access</type>
            <source>Data/DataBase/# Data.asp</source>
        </DB>
        ...
    </SEConfigs>
>>
>
    <%
    SE.module("DB").open()
>
    ' 查询数据
    Dim rs
    Set rs = SE.module("DB").executeSql("SELECT userName FROM UserLists")
>
    SE.module("DB").close()
    %>