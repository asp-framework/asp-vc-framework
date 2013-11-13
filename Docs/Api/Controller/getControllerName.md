getControllerName
=================
`getControllerName` &mdash; 获取当前控制器名称

说明
----
>     string getControllerName(void)
> 获取当前控制器名称。

参数
----
> 没有参数。

返回值
------
> **类型：**`string`  
> **说明：**当前控制器名称。

范例
----
>
    <!-- 配置文件 -->
    <?xml version="1.0" encoding="UTF-8"?>
    <SEConfigs>
        ...
        <Router>
            <appName>HelloWorld</appName>
            <controllerName>Index</controllerName>
            <actionName>index</actionName>
        </Router>
        ...
    </SEConfigs>
>>
>
    <%
    Dim controllerName
    controllerName = SE.module("Controller").getControllerName
    Response.Write(controllerName)
    %>
> 以上内容输出：
>
    Index