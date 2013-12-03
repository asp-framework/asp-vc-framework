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

    Public Function TestCaseNames()
        TestCaseNames = Array( _
            "getUrlTest", _
            "getUrlWithTest" _
        )
    End Function

    Public Sub SetUp()
        SE.getSimpleExtensionsBaseClass.loadConfigs("./ProjectTest/Configs/config.xml")
    End Sub

    Public Sub TearDown()
        'Response.Write("TearDown<br>")
    End Sub

    ' 获取URL测试
    Public Sub getUrlTest(oTestResult)
        vActual = SE.module("Request").getUrl("Dir")
        oTestResult.AssertEquals _
            "/Test/", _
            vActual, _
            "获取URL异常"
        vActual = SE.module("Request").getUrl(0)
        oTestResult.AssertEquals _
            "/Test/", _
            vActual, _
            "获取URL异常"

        vActual = SE.module("Request").getUrl("Path")
        oTestResult.AssertEquals _
            "/Test/TestASPUnit.asp", _
            vActual, _
            "获取URL异常"
        vActual = SE.module("Request").getUrl(1)
        oTestResult.AssertEquals _
            "/Test/TestASPUnit.asp", _
            vActual, _
            "获取URL异常"

        vActual = SE.module("Request").getUrl("DirWith")
        oTestResult.AssertEquals _
            "/Test/?UnitRunner=results", _
            vActual, _
            "获取URL异常"
        vActual = SE.module("Request").getUrl(2)
        oTestResult.AssertEquals _
            "/Test/?UnitRunner=results", _
            vActual, _
            "获取URL异常"

        vActual = SE.module("Request").getUrl("PathWith")
        oTestResult.AssertEquals _
            "/Test/TestASPUnit.asp?UnitRunner=results", _
            vActual, _
            "获取URL异常"
        vActual = SE.module("Request").getUrl(3)
        oTestResult.AssertEquals _
            "/Test/TestASPUnit.asp?UnitRunner=results", _
            vActual, _
            "获取URL异常"
    End Sub

    Public Sub getUrlWithTest(oTestResult)
        vActual = SE.module("Request").getUrlWith("Dir", "a=b")
        oTestResult.AssertEquals _
            "/Test/?a=b", _
            vActual, _
            "获取URL异常"
        vActual = SE.module("Request").getUrlWith(0, "a=b")
        oTestResult.AssertEquals _
            "/Test/?a=b", _
            vActual, _
            "获取URL异常"

        vActual = SE.module("Request").getUrlWith("Path", "a=b")
        oTestResult.AssertEquals _
            "/Test/TestASPUnit.asp?a=b", _
            vActual, _
            "获取URL异常"
        vActual = SE.module("Request").getUrlWith(1, "a=b")
        oTestResult.AssertEquals _
            "/Test/TestASPUnit.asp?a=b", _
            vActual, _
            "获取URL异常"

        vActual = SE.module("Request").getUrlWith("DirWith", "-UnitRunner&a=b")
        oTestResult.AssertEquals _
            "/Test/?a=b", _
            vActual, _
            "获取URL异常"
        vActual = SE.module("Request").getUrlWith(2, "-UnitRunner&a=b")
        oTestResult.AssertEquals _
            "/Test/?a=b", _
            vActual, _
            "获取URL异常"

        vActual = SE.module("Request").getUrlWith("PathWith", "-UnitRunner&a=b")
        oTestResult.AssertEquals _
            "/Test/TestASPUnit.asp?a=b", _
            vActual, _
            "获取URL异常"
        vActual = SE.module("Request").getUrlWith(3, "-UnitRunner&a=b")
        oTestResult.AssertEquals _
            "/Test/TestASPUnit.asp?a=b", _
            vActual, _
            "获取URL异常"
    End Sub

End Class
%>