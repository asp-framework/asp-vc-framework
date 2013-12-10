<%
'''
 ' SimpleExtensionsStringTest.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.12.10
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsStringTest

    Private vActual

    Public Function TestCaseNames()
        TestCaseNames = Array( _
            "md5Test", _
            "htmlFilterTest" _
        )
    End Function

    Public Sub SetUp()
        SE.getSimpleExtensionsBaseClass.loadConfigs("./ProjectTest/Configs/config.xml")
    End Sub

    Public Sub TearDown()
        'Response.Write("TearDown<br>")
    End Sub

    ' MD5 加密测试
    Public Sub md5Test(oTestResult)
        vActual = SE.module("String").md5("SE")

        oTestResult.AssertEquals _
            "f003c44deab679aa2edfaff864c77402", _
            vActual, _
            "导入模块异常"
    End Sub

    ' HTML过滤测试
    Public Sub htmlFilterTest(oTestResult)
        vActual = SE.module("String").htmlFilter( _
            "<div style="">S</div>" & _
            "<span id=""se"">E</span>" & _
            "<img />" _
        )

        oTestResult.AssertEquals _
            "SE", _
            vActual, _
            "HTML过滤异常"
    End Sub

End Class
%>