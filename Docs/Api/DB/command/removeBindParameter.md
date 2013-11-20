removeBindParameter
===================
`removeBindParameter` &mdash; 移除绑定参数

说明
----
>     boolean removeBindParameter(string name)
> 移除绑定参数。

参数
----
> `name`
>> **类型：**`string`  
>> **说明：**绑定的参数名称。

返回值
------
> **类型：**`boolean`  
> **说明：**是否移除成功。

范例
----
>
    <%
    Call SE.module("DB").command.bindParameter(":userName", "Admin", "dbString")
    Call SE.module("DB").command.bindParameter(":id", 1, "dbInteger")
>
    SE.module("DB").command.removeBindParameter(":userName")
    SE.module("DB").command.removeBindParameter(":id")
    %>