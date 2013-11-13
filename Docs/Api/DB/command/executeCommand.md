executeCommand
==============
`executeCommand` &mdash; 执行命令

说明
----
>     recordset executeCommand(void)
> 执行命令

参数
----
> 没有参数。

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
    ' 创建命令
    Dim commandString
    commandString = _
        "SELECT userName " & _
        "FROM UserLists " & _
        "WHERE " & _
            "userName = :userName " & _
            "AND " & _
            "id = :id""
    SE.module("DB").command.createCommand(commandString)
>
    ' 绑定参数
    Call SE.module("DB").command.bindParameter(":userName", "Admin", "dbString")
    Call SE.module("DB").command.bindParameter(":id", 1, "dbInteger")
>
    SE.module("DB").open()
>
    ' 执行命令
    Dim rs
    Set rs = SE.module("DB").command.executeCommand()
>
    SE.module("DB").close()
    %>