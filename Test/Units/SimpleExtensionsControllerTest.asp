<%
'''
 ' SimpleExtensionsControllerTest.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.11.1
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsControllerTest

    Private vActual

    Public Function TestCaseNames()
        TestCaseNames = Array( _
            "getLayoutPathTest", _
            "getViewPathTest", _
            "getControllersDirTest", _
            "getViewsDirTest", _
            "getAppDirTest", _
            "getControllerNameTest" _
        )
    End Function

    Public Sub SetUp()
        SE.getSimpleExtensionsBaseClass.loadConfigs("./ProjectTest/Configs/config.xml")
        SE.module("Router").run()
    End Sub

    Public Sub TearDown()
        'Response.Write("TearDown<br>")
    End Sub

    ' 获取布局路径测试
    Public Sub getLayoutPathTest(oTestResult)
        vActual = SE.module("Controller").getLayoutPath("layout")

        oTestResult.AssertEquals _
            "ProjectTest/AppTest/Test/Views/Layouts/layout.asp", _
            vActual, _
            "获取布局路径异常"
    End Sub

    ' 获取视图路径测试
    Public Sub getViewPathTest(oTestResult)
        vActual = SE.module("Controller").getViewPath("index")

        oTestResult.AssertEquals _
            "ProjectTest/AppTest/Test/Views/Index/index.asp", _
            vActual, _
            "获取视图路径异常"
    End Sub

    ' 获取当前应用控制器目录测试
    Public Sub getControllersDirTest(oTestResult)
        vActual = SE.module("Controller").getControllersDir

        oTestResult.AssertEquals _
            "ProjectTest/AppTest/Test/Controllers", _
            vActual, _
            "获取当前应用控制器目录异常"
    End Sub

    ' 获取当前控制器视图目录测试
    Public Sub getViewsDirTest(oTestResult)
        vActual = SE.module("Controller").getViewsDir

        oTestResult.AssertEquals _
            "ProjectTest/AppTest/Test/Views/Index", _
            vActual, _
            "获取当前控制器视图目录异常"
    End Sub

    ' 获取当前应用目录测试
    Public Sub getAppDirTest(oTestResult)
        vActual = SE.module("Controller").getAppDir

        oTestResult.AssertEquals _
            "ProjectTest/AppTest/Test", _
            vActual, _
            "获取当前应用目录异常"
    End Sub

    ' 获取当前控制器名称测试
    Public Sub getControllerNameTest(oTestResult)
        vActual = SE.module("Controller").getControllerName

        oTestResult.AssertEquals _
            "Index", _
            vActual, _
            "获取当前应用目录异常"
    End Sub

End Class
%>