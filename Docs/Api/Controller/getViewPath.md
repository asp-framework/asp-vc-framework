getViewPath
===========
`getViewPath` &mdash; 获取当前控制器视图路径

说明
----
>     string getViewPath(string viewName)
> 获取当前控制器视图路径。

参数
----
> `viewName`
>> **类型：**`string`  
>> **说明：**视图名称。

返回值
------
> **类型：**`string`  
> **说明：**视图路径。

范例
----
>
    <!-- 配置文件 -->
    <?xml version="1.0" encoding="UTF-8"?>
    <SEConfigs>
        <System>
            <development>True</development>
            <seDir>../Framework</seDir>
            <appsDir>Apps</appsDir>
        </System>
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
    Dim viewPath
    viewPath = SE.module("Controller").getViewPath("index")
    Response.Write(viewPath)
    %>
> 以上内容输出：
>
    Apps/HelloWorld/Views/Index/index.asp