getLayoutPath
=============
`getLayoutPath` &mdash; 获取布局路径

说明
----
>     string getLayoutPath(ByVal layoutName)
> 获取布局路径。

参数
----
> `layoutName`
>> **类型：**`string`   
>> **说明：**布局名称。  
>> **范例：**`"layout"`

返回值
------
> **类型：**`string`  
> **说明：**布局路径。(路径格式为相对路径，相对路径起始于 **应用目录** 的设置。)

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
    Dim layoutPath
    layoutPath = SE.module("Controller").getLayoutPath("layout")
    Response.Write(layoutPath)
    %>
>   以上内容输出:
>
    Apps/HelloWorld/Views/Layouts/layout.asp