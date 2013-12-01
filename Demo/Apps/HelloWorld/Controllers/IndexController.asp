<%
'''
 ' 首页
 ''
%>

<%
Class IndexController

    Public Sub indexAction()
        Dim parameters
        Set parameters = Server.CreateObject("Scripting.Dictionary")
        Call parameters.Add("title", "SE")
        Call parameters.Add("content", "Hello World")

SE.module("Debugging").enabled()
SE.module("Debugging").renderPanel()

Response.Write(SE.module("Request").getUrlWith("DirWith", "-k"))
Response.Write("<br />")

        Call SE.module("View").render( _
            "index", _
            "layout", _
            parameters _
        )
    End Sub

End Class
%>