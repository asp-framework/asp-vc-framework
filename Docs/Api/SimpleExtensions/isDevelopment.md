isDevelopment
=============
`isDevelopment` &mdash; 判断是否开发环境

说明
----
>     boolean isDevelopment(void)
> 判断是否开发环境

参数
----
> 没有参数。

返回值
------
> **类型：**`boolean`  
> **说明：**是否开发环境。(默认值为`False`)

范例
----
>
    <%
    Dim development
    development = SE.isDevelopment
    If development Then
        Response.Write("当前为开发环境。")
    Else
        Response.Write("当前为生产环境。")
    End If
    %>