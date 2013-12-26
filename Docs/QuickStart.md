快速入门
========

文件目录
--------
```
Demo/                                           实例目录
    Apps/                                       应用目录
        HelloWorld/                             "HelloWorld"应用
        Controllers/                            控制器目录
            IndexController.asp                 "Index"控制器
        Views/                                  视图目录
            Index/                              "Index"控制器视图目录
                index.asp                       "index"视图
            Layouts/                            布局目录
                layout.asp                      "layout"布局
    Configs/                                    配置文件目录
        config.xml                              框架配置文件
    index.asp                                   入口文件
Framework/                                      核心框架目录
    Controller/                                 控制器模块
    DB/                                         数据库模块
    Debugging/                                  调试模块
    Error/                                      错误异常模块
    File/                                       文件模块
    I18N/                                       国际化模块
    Request/                                    请求模块
    Router/                                     路由器模块
    String/                                     字符串处理模块
    View/                                       视图渲染模块
    SimpleExtensions.asp                        SE框架类
    SimpleExtensionsBase.asp                    SE框架基类
```

入口文件
--------
`Demo/index.asp`

站点入口文件。导入 `SimpleExtensions` 类文件，然后调用 `run()` 函数启动框架。  
**( `run()` 函数需要传入配置文件路径以配置启动框架。 )**

```html5
<%@
    Language = "VBScript"
    CodePage = "65001"
%>
<% Option Explicit %>
<%
    Session.CodePage = 65001
    Response.Charset = "UTF-8"

    Response.CacheControl = "no-cache"
    Call Response.AddHeader("Pragma", "no-cache")
    Response.Expires = -1
%>

<!-- #include file = "../Framework/SimpleExtensions.asp" -->

<% SE.run("Configs/config.xml") %>
```

配置文件
--------
`Demo/Comfigs/config.xml`

配置文件为XML文件，配置项都包含在 `SEConfigs` 标签内。  
`SEConfigs` 标签为系统模块，`seDir`(框架目录) 、 `appsDir`(应用目录) 为必须配置项。  
`Router` 标签为路由器模块，`appName`(应用) 、 `controllerName`(控制器) 、 `actionName`(动作) 为必须设置的默认值。

```xml
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
    <I18N>
        <language>zh-cn</language>
    </I18N>
    <DB>
        <type>access</type>
        <source>../Test/ProjectTest/Data/DataBase/# Data.asp</source>
    </DB>
    <Error>
        <redirectURL>./</redirectURL>
    </Error>
</SEConfigs>
```

控制器
------
`Demo/Apps/HelloWorld/Controllers/IndexController.asp`

控制器类以大驼峰方式命名，命名规则为 `控制器名` + `Controller` 后缀。  
控制器文件名与类名保持一致。  
下面的控制器调用 `Render`(视图渲染) 模块，调用了 `layout` 布局 和 `index` 视图，同时传入 `parameters` 参数。  
**(`Render`(视图渲染) 模块 `rendering` 方法 传入的 `parameters` 变量必须为 `Dictionary` 类型)**

```html5
<%
'''
 ' 首页控制器
 ''
%>

<%
Class IndexController

    Public Sub indexAction()
        Dim parameters
        Set parameters = Server.CreateObject("Scripting.Dictionary")
        Call parameters.Add("title", "SE")
        Call parameters.Add("helloWorldText", "Hello, World!")

        Call SE.module("View").render( _
            "index", _
            "layout", _
            parameters _
        )
    End Sub

End Class
%>
```

布局
----
`Demo/Apps/HelloWorld/Views/Layouts/layout.asp`

布局文件中使用 `<% '<!-- #content -->' %>` 标签即可调用对应的视图内容。  
`<% '<!-- #contentStartToDo -->' %>` 标签将调用视图中 `<% '<!-- #contentStart -->' %>` 标签之后的内容。  
`<% '<!-- #contentEndToDo -->' %>` 标签将调用视图中 `<% '<!-- #contentEnd -->' %>` 标签之后的内容。

```html5
<% Response.Buffer = True %>
<!DOCTYPE html>
<html>
<head>
    <title><%= title %></title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <link rel="stylesheet" href="http://cdn.bootcss.com/twitter-bootstrap/3.0.3/css/bootstrap.min.css" />
    <% '<!-- #contentStartToDo -->' %>
</head>
<% Response.Flush() %>
<body>

    <!-- 导航条 -->
    <nav class="navbar navbar-inverse" role="navigation">
        <div class="container">
            <div class="navbar-header">
                <a class="navbar-brand" href="./">SE&nbsp;For&nbsp;ASP</a>
            </div>
            <div class="collapse navbar-collapse">

            </div>
        </div>
    </nav>
    <!-- /导航条 -->

    <% '<!-- #content -->' %>

    <% Response.Flush() %>
    <!-- Js -->
    <script src="http://cdn.bootcss.com/jquery/1.10.2/jquery.min.js"></script>
    <script src="http://cdn.bootcss.com/twitter-bootstrap/3.0.3/js/bootstrap.min.js"></script>
    <% '<!-- #contentEndToDo -->' %>
    <!-- /Js -->

</body>
</html>
```

视图
----
`Demo/Apps/HelloWorld/Views/Index/index.asp`

`content` 为控制器中传入的变量。

```html5
<style>
    /* 样式 */
</style>
<% '<!-- #contentStart -->' %>

<div class="jumbotron">
    <div class="container">
        <h1><%= helloWorldText %></h1>
        <p>SE&nbsp;For&nbsp;ASP&nbsp;是一个以Vbscript语言为基础的ASP&nbsp;VB开发框架。</p>
    </div>
</div>

<% '<!-- #contentEnd -->' %>
<script>
    // Js脚本 
</script>
```