<%
'''
 ' SimpleExtensionsRouterTest.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.10.30
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsRouterTest

    Private vActual

    Public Function TestCaseNames()
        TestCaseNames = Array( _
            "getAppNameTest", _
            "getControllerNameTest", _
            "getActionNameTest" _
        )
    End Function

    Public Sub SetUp()
        SE.getSimpleExtensionsBaseClass.loadConfigs("./ProjectTest/Configs/config.xml")
        SE.module("Router").run()
    End Sub

    Public Sub TearDown()
        'Response.Write("TearDown<br>")
    End Sub

    ' 获取应用名称测试
    Public Sub getAppNameTest(oTestResult)
        vActual = SE.module("Router").getAppName

        oTestResult.AssertEquals _
            "Test", _
            vActual, _
            "获取应用名称异常"
    End Sub

    ' 获取控制器名称测试
    Public Sub getControllerNameTest(oTestResult)
        vActual = SE.module("Router").getControllerName

        oTestResult.AssertEquals _
            "Index", _
            vActual, _
            "获取控制器名称异常"
    End Sub

    ' 获取动作名称测试
    Public Sub getActionNameTest(oTestResult)
        vActual = SE.module("Router").getActionName

        oTestResult.AssertEquals _
            "index", _
            vActual, _
            "获取动作名称异常"
    End Sub

End Class
%>