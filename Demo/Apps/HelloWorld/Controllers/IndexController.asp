<%
'''
 ' 首页控制器
 ''
%>

<%
Class IndexController

    Public Sub indexAction()
        Dim parameters
        Set parameters = Server.CreateObject("Scripting.Dictionary")
        Call parameters.Add("title", "SE")
        Call parameters.Add("helloWorldText", "Hello, World!")

        Call SE.module("View").render( _
            "index", _
            "layout", _
            parameters _
        )
    End Sub

End Class
%>