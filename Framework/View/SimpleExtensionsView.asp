<%
'''
 ' SimpleExtensionsRender.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.11.12
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsView

    ' @var string <布局可执行代码>
    Private layoutCode

    ' @var string <内容可执行代码>
    Private contentCode

    ' @var string <内容结束后要执行的代码>
    Private contentEndToDoCode

'###########################'
'###########################'

    '''
     ' 渲染视图
     '
     ' @param string viewName <视图名称>
     ' @param string|null layoutName <布局名称>
     ' @param dictionary|null &parameters <参数>
     ''
    Public Function render(ByVal viewName, ByVal layoutName, ByRef parameters)
        contentCode = SE.getIncludeCode(SE.module("Controller").getViewPath(viewName))
        separateContentCode()

        ' 定义传入变量
        If Not IsNull(parameters) Then
            For Each key In parameters.Keys
                Execute("Dim " & key)
                Execute(key & " = parameters.Item(key)")
            Next
        End If

        ' 渲染视图
        If IsNull(layoutName) Then
            Execute(contentCode)
            Exit Function
        End If

        ' 渲染布局
        layoutCode = SE.getIncludeCode(SE.module("Controller").getLayoutPath(layoutName))
        replaceSETag()
        Execute(layoutCode)
    End Function

    '''
     ' 按 '<!-- #contentEnd -->' 分离视图代码
     ''
    Private Function separateContentCode()
        Dim CONTENTEND_TAG_LEFT, CONTENTEND_TAG, CONTENTEND_TAG_RIGHT
        CONTENTEND_TAG_LEFT = "'<!--"
        CONTENTEND_TAG = "#contentEnd"
        CONTENTEND_TAG_RIGHT = "-->'"

        Dim tagStart, tagEnd
        tagStart = InStr(1, contentCode, CONTENTEND_TAG_LEFT) + 5
        tagEnd = InStr(tagStart, contentCode, CONTENTEND_TAG_RIGHT) + 4
        If InStr(1, Trim(Mid(contentCode, tagStart, tagEnd - tagStart - 4)), CONTENTEND_TAG, 1) = 1 Then
            contentEndToDoCode = Mid(contentCode, tagEnd)
            contentCode = Mid(contentCode, 1, tagStart - 6)
        End If
    End Function

    '''
     ' 替换布局中的特殊标签
     ''
    Private Function replaceSETag()
        replaceContentTag()
        replaceContentEndToDoTag()
    End Function

    '''
     ' 替换布局中的 '<!-- #content -->' 标签
     ''
    Private Function replaceContentTag()
        Dim CONTENT_TAG_LEFT, CONTENT_TAG, CONTENT_TAG_RIGHT
        CONTENT_TAG_LEFT = "'<!--"
        CONTENT_TAG = "#content"
        CONTENT_TAG_RIGHT = "-->'"

        Dim tagStart, tagEnd
        tagStart = InStr(1, layoutCode, CONTENT_TAG_LEFT) + 5
        tagEnd = InStr(tagStart, layoutCode, CONTENT_TAG_RIGHT) + 4
        If InStr(1, Trim(Mid(layoutCode, tagStart, tagEnd - tagStart - 4)), CONTENT_TAG, 1) = 1 Then _
            layoutCode = Mid(layoutCode, 1, tagStart - 6) & contentCode & Mid(layoutCode, tagEnd)
    End Function

    '''
     ' 替换布局中的 '<!-- #contentEndToDo -->' 标签
     ''
    Private Function replaceContentEndToDoTag()
        Dim CONTENTENDTODO_TAG_LEFT, CONTENTENDTODO_TAG, CONTENTENDTODO_TAG_RIGHT
        CONTENTENDTODO_TAG_LEFT = "'<!--"
        CONTENTENDTODO_TAG = "#contentEndToDO"
        CONTENTENDTODO_TAG_RIGHT = "-->'"

        Dim tagStart, tagEnd
        tagStart = InStr(1, layoutCode, CONTENTENDTODO_TAG_LEFT) + 5
        tagEnd = InStr(tagStart, layoutCode, CONTENTENDTODO_TAG_RIGHT) + 4
        If InStr(1, Trim(Mid(layoutCode, tagStart, tagEnd - tagStart - 4)), CONTENTENDTODO_TAG, 1) = 1 Then _
            layoutCode = Mid(layoutCode, 1, tagStart - 6) & contentEndToDoCode & Mid(layoutCode, tagEnd)
    End Function

End Class
%>