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

        ' 定义传入变量
        If Not IsNull(parameters) Then
            For Each key In parameters.Keys
                Execute("Dim " & key)
                Execute(key & " = parameters.Item(key)")
            Next
        End If

        If IsNull(layoutName) Then
            Execute(contentCode)
        Else
            Execute( _
                replaceLayoutContentTag( _
                    SE.getIncludeCode(SE.module("Controller").getLayoutPath(layoutName)) _
                ) _
            )
        End If
    End Function

    '''
     ' 替换布局视图标签
     ''
    Private Function replaceLayoutContentTag(ByVal layoutCode)
        Dim CONTENT_TAG_LEFT, CONTENT_TAG, CONTENT_TAG_RIGHT
        CONTENT_TAG_LEFT = "'<!--" : CONTENT_TAG = "#content" : CONTENT_TAG_RIGHT = "-->'"

        Dim tagStart, tagEnd
        tagStart = InStr(1, layoutCode, CONTENT_TAG_LEFT) + 5
        tagEnd = InStr(tagStart, layoutCode, CONTENT_TAG_RIGHT) + 4
        If InStr(1, Trim(Mid(layoutCode, tagStart, tagEnd - tagStart - 4)), "#content", 1) = 1 Then
            layoutCode = _
                Mid(layoutCode, 1, tagStart - 6) _
                & vbCrLf & contentCode & vbCrLf _
                & Mid(layoutCode, tagEnd)
        End If

        replaceLayoutContentTag = layoutCode
    End Function

End Class
%>