getDBType
=========
`getDBType` &mdash; 获取数据库类型

说明
----
>     string getDBType(void)
> 获取数据库类型。

参数
----
> 没有参数。

返回值
------
> **类型：**`string`  
> **说明：**数据库类型。
>> ACCESS：`"access"`。

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
    Dim dbType
    dbType = SE.module("DB").getDBType
    Response.Write(dbType)
    %>
> 以上内容输出：
>
    access