removeAllBindParameter
======================
`removeAllBindParameter` &mdash; 移除所有绑定参数

说明
----
>     boolean removeAllBindParameter(void)
> 移除所有绑定参数。

参数
----
> 没有参数。

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
    SE.module("DB").command.removeAllBindParameter()
    %>