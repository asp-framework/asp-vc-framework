runAction
=========
`runAction` &mdash; 运行动作

说明
----
>     void runAction(string controllerName, string actionName)
> 运行动作。

参数
----
> `controllerName`
>> **类型：**`string`  
>> **说明：**控制器名称。
>
> `actionName`
>> **类型：**`string`  
>> **说明：**动作名称。

返回值
------
> 没有返回值。

范例
----
>
    <%
    Class IndexController
        Public Sub indexAction()
            Response.Write("运行动作。")
        End Sub
    End Class
    %>
>>
>
    <%
    SE.module("Controller").runAction("Index", "index")
    %>
> 以上内容输出：
>
    运行动作。