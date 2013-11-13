getDBName
=========
`getDBName` &mdash; 获取数据库名称

说明
----
>     string|empty getDBName(void)
> 获取数据库名称。(部分数据库类型没有设置数据库名称。)

参数
----
> 没有参数。

返回值
------
> **类型：**`string` | `empty`  
> **说明：**
>> 已设置 **数据库名称** 返回`string`；  
>> 没有设置 **数据库名称** 返回`empty`。

范例
----
>
    <!-- 配置文件 -->
    <?xml version="1.0" encoding="UTF-8"?>
    <SEConfigs>
        ...
        <DB>
            ...
            <dbName>SE</dbName>
            ...
        </DB>
        ...
    </SEConfigs>
>>
>
    <%
    Dim dbName
    dbName = SE.module("DB").getDBName
    Response.Write(dbName)
    %>
> 以上内容输出：
>
    SE