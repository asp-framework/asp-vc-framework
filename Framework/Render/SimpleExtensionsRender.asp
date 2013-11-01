<%
'''
 ' SimpleExtensionsRender.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.11.1
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsRender

    ' @var string <内容可执行代码>
    Private contentCode

'###########################'
'###########################'

    '''
     ' 渲染视图
     '
     ' @param string viewName <视图名称>
     ' @param string|null layoutName <布局名称>
     ' @param dictionary|null &parameters <参数>
     ''
    Public Function rendering(ByVal viewName, ByVal layoutName, ByRef parameters)
        contentCode = SE.getIncludeCode(SE.module("Controller").getViewPath(viewName))
        If IsNull(layoutName) Then
            Execute(contentCode)
        Else
            Execute(SE.getIncludeCode(SE.module("Controller").getLayoutPath(layoutName)))
        End If
    End Function

    '''
     ' 执行内容
     ''
    Private Sub content()
        Execute(contentCode)
    End Sub

End Class
%>