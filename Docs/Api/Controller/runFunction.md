runFunction
===========
`runFunction` &mdash; 运行方法

说明
----
>     void runFunction(string controllerName, string functionName, array parameters)
> 运行方法。

参数
----
> `controllerName`
>> **类型：**`string`  
>> **说明：**控制器名称。
>
> `functionName`
>> **类型：**`string`  
>> **说明：**方法名称。
>
> `parameters`
>> **类型：**`array`  
>> **说明：**方法需要的参数。

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
        Public Function test(ByVal testString)
            Response.Write(testString)
        End Function
    End Class
    %>
>>
>
    <%
    Call SE.module("Controller").runFunction("Index", "index", Null)
    Call SE.module("Controller").runFunction("Index", "test", Array("运行方法。"))
    %>
> 以上内容输出：
>
    运行动作。
    运行方法。