getAppDir
=========
`getAppDir` &mdash; 获取当前应用目录

说明
----
>     string getAppDir(void)
> 获取当前应用目录。

参数
----
> 没有参数。

返回值
------
> **类型：**`string`  
> **说明：**当前应用目录。

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
    Dim appDir
    appDir = SE.module("Controller").getAppDir
    Response.Write(appDir)
    %>
> 以上内容输出：
>
    Apps/HelloWorld