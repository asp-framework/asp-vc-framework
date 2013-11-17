bindParameter
=============
`bindParameter` &mdash; 绑定参数

说明
----
>     boolean bindParameter(string name, string value, string dataType)
> 绑定参数。(创建命令时使用`":参数名"`调用绑定参数。)

参数
----
> `name`
>> **类型：**`string`  
>> **说明：**绑定的参数名。

> `value`
>> **类型：**`string`  
>> **说明：**绑定的参数值。

> `dataType`
>> **类型：**`string`  
>> **说明：**绑定参数的类型。  
>>> 字符串：`"dbString"`；  
>>> 整形：`"dbInteger"`。

返回值
------
> **类型：**`boolean`  
> **说明：**是否绑定成功。

范例
----
>
    <%
    Call SE.module("DB").command.bindParameter(":userName", "Admin", "dbString")
    Call SE.module("DB").command.bindParameter(":id", 1, "dbInteger")
    %>