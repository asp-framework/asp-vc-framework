<%
'''
 ' 首页
 ''
%>

<%
Class IndexController

    Public Sub indexAction()
        Call SE.module("Render").rendering( _
            "index", _
            "layout", _
            Null _
        )
    End Sub

End Class
%>