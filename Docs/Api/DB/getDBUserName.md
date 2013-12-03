getDBUserName
=============
`getDBUserName` &mdash; 获取数据库用户名

说明
----
>     string|empty getDBUserName(void)
> 获取数据库用户名。(部分数据库类型没有设置数据库用户名。)

参数
----
> 没有参数。

返回值
------
> **类型：**`string` | `empty`  
> **说明：**数据库用户名。
>> 已设置 **数据库用户名** 返回`string`；  
>> 没有设置 **数据库用户名** 返回`empty`。

范例
----
>
    <!-- 配置文件 -->
    <?xml version="1.0" encoding="UTF-8"?>
    <SEConfigs>
        ...
        <DB>
            ...
            <userName>SE</userName>
            ...
        </DB>
        ...
    </SEConfigs>
>>
>
    <%
    Dim userName
    userName = SE.module("DB").getDBUserName
    Response.Write(userName)
    %>
> 以上内容输出：
>
    SE