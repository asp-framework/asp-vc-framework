<%
'''
 ' SimpleExtensionsErrorTest.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.11.6
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsErrorTest

    Private vActual

    Public Function TestCaseNames()
        TestCaseNames = Array( _
            "errorTest" _
        )
    End Function

    Public Sub SetUp()
        ' Response.Write("SetUp<br>")
    End Sub

    Public Sub TearDown()
        'Response.Write("TearDown<br>")
    End Sub

    ' 错误测试
    Public Sub errorTest(oTestResult)
        vActual = SE.module("Error").getErrorDefine(0)
SE.module("Error").throwError(0)
        oTestResult.AssertEquals _
            "读取文件测试", _
            vActual, _
            "错误异常"
    End Sub

End Class
%>