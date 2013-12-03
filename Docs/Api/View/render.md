render
======
`render` &mdash; 渲染视图

说明
----
>     void render(string viewName, string|null layoutName, dictionary|null parameters)
> 渲染视图。

参数
----
> `viewName`
>> **类型：**`string`  
>> **说明：**视图名称。
>
> `layoutName`
>> **类型：**`string` | `null`  
>> **说明：**布局名称。(值为`null`时只渲染视图。)
>
> `parameters`
>> **类型：**`dictionary` | `null`  
>> **说明：**参数。

返回值
------
> 没有返回值。

范例
----
>
    <%
    Class IndexController
        Public Sub indexAction()
            Dim parameters
            Set parameters = Server.CreateObject("Scripting.Dictionary")
            Call parameters.Add("title", "SE")
            Call parameters.Add("content", "Hello World")
>
            Call SE.module("View").render( _
                "index", _
                "layout", _
                parameters _
            )
        End Sub
    End Class
    %>