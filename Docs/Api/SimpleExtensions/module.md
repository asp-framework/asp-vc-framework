module
======
`module` &mdash; 调用模块

说明
----
>     class module(string moduleName)
> 调用模块。

参数
----
> `moduleName`
>> **类型：**`string`  
>> **说明：**模块名称。

返回值
------
> **类型：**`class`  
> **说明：**模块类。

范例
----
>
    <%
    Dim md5String
    md5String = SE.module("String").md5("SE")
    Response.Write(md5String)
    %>
> 上面内容调用了`String`模块，并执行了其中的`md5`方法，将输出如下内容：
>
    f003c44deab679aa2edfaff864c77402