<%
'''
 ' SimpleExtensionsFileTest.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.11.7
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsFileTest

    Private vActual

    Public Function TestCaseNames()
        TestCaseNames = Array( _
            "fileExistsTest", _
            "dirExistsTest" _
        )
    End Function

    Public Sub SetUp()
        ' Response.Write("SetUp<br>")
    End Sub

    Public Sub TearDown()
        'Response.Write("TearDown<br>")
    End Sub

    ' 判断文件是否存在测试
    Public Sub fileExistsTest(oTestResult)
        vActual = SE.module("File").fileExists("./ProjectTest/IncludeTest/loadFileTest.asp")

        oTestResult.AssertEquals _
            True, _
            vActual, _
            "文件判断异常"
    End Sub

    ' 判断文件是否存在测试
    Public Sub dirExistsTest(oTestResult)
        vActual = SE.module("File").dirExists("./ProjectTest/IncludeTest")

        oTestResult.AssertEquals _
            True, _
            vActual, _
            "文件判断异常"
    End Sub

End Class
%>