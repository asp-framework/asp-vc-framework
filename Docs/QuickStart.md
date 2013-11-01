快速入门
========

文件目录
--------

    Demo/                                   实例目录
        Apps/                               应用目录
            HelloWorld/                     "HelloWorld"应用
            Controllers/                    控制器目录
                IndexController.asp         "Index"控制器
            Views/                          视图目录
                Index/                      "Index"控制器视图目录
                    index.asp               "index"视图
                Layouts/                    布局目录
                    layout.asp              "layout"布局
        Configs/                            配置文件目录
            configs.xml                     框架配置文件
        index.asp                           入口文件
    Framework/                              核心框架目录

### 入口文件

站点入口文件。

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

<%
    SE.run("Configs/config.xml")
%>
~~~