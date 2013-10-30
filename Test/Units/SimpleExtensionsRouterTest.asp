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
            "loadFileTest" _
        )
    End Function

    Public Sub SetUp()
        Set SimpleExtensionsRouterClass = New SimpleExtensionsRouter
    End Sub

    Public Sub TearDown()
        'Response.Write("TearDown<br>")
    End Sub

End Class
%>