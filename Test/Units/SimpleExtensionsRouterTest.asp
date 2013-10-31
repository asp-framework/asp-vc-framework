<%
'''
 ' SimpleExtensionsRouterTest.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.10.30
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<!-- #include file = "../../Framework/Router/SimpleExtensionsRouter.asp" -->

<%
Class SimpleExtensionsRouterTest

    Private SimpleExtensionsRouterClass

    Private vActual

    Public Function TestCaseNames()
        TestCaseNames = Array( _
            "getAppNameTest", _
            "getControllerNameTest", _
            "getActionNameTest" _
        )
    End Function

    Public Sub SetUp()
        Set SimpleExtensionsRouterClass = New SimpleExtensionsRouter
    End Sub

    Public Sub TearDown()
        'Response.Write("TearDown<br>")
    End Sub

    ' 获取应用名称测试
    Public Sub getAppNameTest(oTestResult)
        SimpleExtensionsRouterClass.run()
        vActual = SimpleExtensionsRouterClass.getAppName

        oTestResult.AssertEquals _
            "HelloWorld", _
            vActual, _
            "读取文件信息异常"
    End Sub

    ' 获取应用名称测试
    Public Sub getControllerNameTest(oTestResult)
        SimpleExtensionsRouterClass.run()
        vActual = SimpleExtensionsRouterClass.getControllerName

        oTestResult.AssertEquals _
            "Index", _
            vActual, _
            "读取文件信息异常"
    End Sub

    ' 获取应用名称测试
    Public Sub getActionNameTest(oTestResult)
        SimpleExtensionsRouterClass.run()
        vActual = SimpleExtensionsRouterClass.getActionName

        oTestResult.AssertEquals _
            "index", _
            vActual, _
            "读取文件信息异常"
    End Sub

End Class
%>