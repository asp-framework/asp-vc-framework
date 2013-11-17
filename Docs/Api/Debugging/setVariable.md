setVariable
===========
`setVariable` &mdash; 设置参数

说明
----
>     void setVariable(string variableName, mix variable)
> 设置参数。(设置的参数会在 **变量** 面板中显示。)

参数
----
> `variableName`
>> **类型：**`string`  
>> **说明：**参数名称。
>
> `variable`
>> **类型：**`mix`  
>> **说明：**参数。

返回值
------
> 没有返回值。

范例
----
>
    <%
    Call SE.module("Debugging").setVariable("module", "Debugging")
>
    Dim module : module = "Debugging"
    Call SE.module("Debugging").setVariable("module", module)
    %>