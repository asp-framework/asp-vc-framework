getDBSource
===========
`getDBSource` &mdash; 获取数据库源

说明
----
>     string|empty getDBSource(void)
> 获取数据库源。(部分数据库类型没有设置数据源。)

参数
----
> 没有参数。

返回值
------
> **类型：**`string` | `empty`  
> **说明：**
>> 已设置数据源返回`string`；  
>> 没有设置数据源返回`empty`。

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
    Dim dbSource
    dbSource = SE.module("DB").getDBSource
    Response.Write(dbSource)
    %>
> 以上内容输出：
>
    Data/DataBase/# Data.asp