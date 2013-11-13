getDBPassword
=============
`getDBPassword` &mdash; 获取数据库密码

说明
----
>     string|empty getDBPassword(void)
> 获取数据库密码。(部分数据库类型没有设置数据库密码。)

参数
----
> 没有参数。

返回值
------
> **类型：**`string` | `empty`  
> **说明：**
>> 已设置 **数据库密码** 返回`string`；  
>> 没有设置 **数据库密码** 返回`empty`。

范例
----
>
    <!-- 配置文件 -->
    <?xml version="1.0" encoding="UTF-8"?>
    <SEConfigs>
        ...
        <DB>
            ...
            <password>SE</password>
            ...
        </DB>
        ...
    </SEConfigs>
>>
>
    <%
    Dim password
    password = SE.module("DB").getDBPassword
    Response.Write(password)
    %>
> 以上内容输出：
>
    SE