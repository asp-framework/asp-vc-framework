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
Response.Write(Request.ServerVariables("REQUEST_METHOD"))
Response.Write("<br />")
SE.module("Request")
        Call SE.module("View").render( _
            "index", _
            "layout", _
            parameters _
        )
    End Sub

End Class
%>