快速入门
========

文件目录
--------

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
            configs.xml                             框架配置文件
        index.asp                                   入口文件
    Framework/                                      核心框架目录
        Controller/                                 控制器模块
            SimpleExtensionsController.asp
        Render/                                     视图渲染模块
            SimpleExtensionsRender.asp
        Router/                                     路由器模块
            SimpleExtensionsRouter.asp
        String/                                     字符串处理模块
            SimpleExtensionsString.asp
            SimpleExtensionsStringMD5.asp
        SimpleExtensions.asp                        SE框架类
        SimpleExtensionsBase.asp                    SE框架基类

### 入口文件

站点入口文件。导入 `SimpleExtensions` 类文件，然后调用 `run()` 函数启动框架。

**( `run()` 函数需要传入配置文件路径以配置启动框架。 )**

~~~
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
~~~

### 配置文件

配置文件为XML文件，配置项都包含在 `SEConfigs` 标签内。

`SEConfigs` 标签为系统模块，`seDir`(框架目录) 、 `appsDir`(应用目录) 为必须配置项。

`Router` 标签为路由器模块，`appName`(应用) 、 `controllerName`(控制器) 、 `actionName`(动作) 为必须设置的默认值

~~~
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
</SEConfigs>
~~~