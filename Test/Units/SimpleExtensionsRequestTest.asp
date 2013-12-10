<%
'''
 ' SimpleExtensionsRequestTest.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.12.3
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsRequestTest

    Private vActual

    Private pathInfo

    Private dirInfo

    Public Function TestCaseNames()
        TestCaseNames = Array( _
            "getUrlTest", _
            "getUrlWithTest" _
        )
    End Function

    Public Sub SetUp()
        SE.getSimpleExtensionsBaseClass.loadConfigs("./ProjectTest/Configs/config.xml")
        pathInfo = Request.ServerVariables("PATH_INFO")
        dirInfo = Left(pathInfo, InStrRev(pathInfo, "/"))
    End Sub

    Public Sub TearDown()
        'Response.Write("TearDown<br>")
    End Sub

    ' 获取URL测试
    Public Sub getUrlTest(oTestResult)
        vActual = SE.module("Request").getUrl("Dir")
        oTestResult.AssertEquals _
            dirInfo, _
            vActual, _
            "获取URL异常"
        vActual = SE.module("Request").getUrl(0)
        oTestResult.AssertEquals _
            dirInfo, _
            vActual, _
            "获取URL异常"

        vActual = SE.module("Request").getUrl("Path")
        oTestResult.AssertEquals _
            pathInfo, _
            vActual, _
            "获取URL异常"
        vActual = SE.module("Request").getUrl(1)
        oTestResult.AssertEquals _
            pathInfo, _
            vActual, _
            "获取URL异常"

        vActual = SE.module("Request").getUrl("DirWith")
        oTestResult.AssertEquals _
            dirInfo & "?UnitRunner=results", _
            vActual, _
            "获取URL异常"
        vActual = SE.module("Request").getUrl(2)
        oTestResult.AssertEquals _
            dirInfo & "?UnitRunner=results", _
            vActual, _
            "获取URL异常"

        vActual = SE.module("Request").getUrl("PathWith")
        oTestResult.AssertEquals _
            pathInfo & "?UnitRunner=results", _
            vActual, _
            "获取URL异常"
        vActual = SE.module("Request").getUrl(3)
        oTestResult.AssertEquals _
            pathInfo & "?UnitRunner=results", _
            vActual, _
            "获取URL异常"
    End Sub

    Public Sub getUrlWithTest(oTestResult)
        vActual = SE.module("Request").getUrlWith("Dir", "a=b")
        oTestResult.AssertEquals _
            dirInfo & "?a=b", _
            vActual, _
            "获取URL异常"
        vActual = SE.module("Request").getUrlWith(0, "a=b")
        oTestResult.AssertEquals _
            dirInfo & "?a=b", _
            vActual, _
            "获取URL异常"

        vActual = SE.module("Request").getUrlWith("Path", "a=b")
        oTestResult.AssertEquals _
            pathInfo & "?a=b", _
            vActual, _
            "获取URL异常"
        vActual = SE.module("Request").getUrlWith(1, "a=b")
        oTestResult.AssertEquals _
            pathInfo & "?a=b", _
            vActual, _
            "获取URL异常"

        vActual = SE.module("Request").getUrlWith("DirWith", "-UnitRunner&a=b")
        oTestResult.AssertEquals _
            dirInfo & "?a=b", _
            vActual, _
            "获取URL异常"
        vActual = SE.module("Request").getUrlWith(2, "-UnitRunner&a=b")
        oTestResult.AssertEquals _
            dirInfo & "?a=b", _
            vActual, _
            "获取URL异常"

        vActual = SE.module("Request").getUrlWith("PathWith", "-UnitRunner&a=b")
        oTestResult.AssertEquals _
            pathInfo & "?a=b", _
            vActual, _
            "获取URL异常"
        vActual = SE.module("Request").getUrlWith(3, "-UnitRunner&a=b")
        oTestResult.AssertEquals _
            pathInfo & "?a=b", _
            vActual, _
            "获取URL异常"
    End Sub

End Class
%>