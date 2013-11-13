createCommand
=============
`createCommand` &mdash; 创建命令

说明
----
>     void createCommand(ByVal sqlString)
> 创建命令。

参数
----
> `sqlString`
>> **类型：**`string`  
>> **说明：**SQL命令字符串。(使用`":参数名"`调用绑定参数。)  
>> **范例：**
>>
    SELECT userName
    FROM UserLists
    WHERE
        userName = :userName
        AND
        id = :id

返回值
------
> 没有返回值。

范例
----
>
    <%
    Dim commandString
    commandString = _
        "SELECT userName " & _
        "FROM UserLists " & _
        "WHERE " & _
            "userName = :userName " & _
            "AND " & _
            "id = :id""
    SE.module("DB").command.createCommand(commandString)
    %>