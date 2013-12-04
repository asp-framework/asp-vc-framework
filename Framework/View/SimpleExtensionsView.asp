<%
'''
 ' SimpleExtensionsRender.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.12.4
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
        contentCode = SE.getIncludeCode( _
            SE.module("Controller").getViewPath(viewName) _
        )
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
        layoutCode = SE.getIncludeCode( _
            SE.module("Controller").getLayoutPath(layoutName) _
        )
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
        If tagStart = 5 Then Exit Function
        tagEnd = InStr(tagStart, contentCode, CONTENTEND_TAG_RIGHT) + 4

        Dim searchedContentEndTag
        searchedContentEndTag = Trim(Mid(contentCode, tagStart, tagEnd - tagStart - 4))
        If InStr(1, searchedContentEndTag, CONTENTEND_TAG, 1) = 1 Then
            contentEndToDoCode = Mid(contentCode, tagEnd+2)
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
        If tagStart = 5 Then Exit Function
        tagEnd = InStr(tagStart, layoutCode, CONTENT_TAG_RIGHT) + 4

        Dim searchedContentTag
        searchedContentTag = Trim(Mid(layoutCode, tagStart, tagEnd - tagStart - 4))
        If InStr(1, searchedContentTag, CONTENT_TAG, 1) = 1 Then _
            layoutCode = Mid(layoutCode, 1, tagStart - 6) & contentCode & _
                Mid(layoutCode, tagEnd)
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
        If tagStart = 5 Then Exit Function
        tagEnd = InStr(tagStart, layoutCode, CONTENTENDTODO_TAG_RIGHT) + 4

        Dim searchedContentendToDoTag
        searchedContentendToDoTag = Trim(Mid(layoutCode, tagStart, tagEnd - tagStart - 4))
        If InStr(1, searchedContentendToDoTag, CONTENTENDTODO_TAG, 1) = 1 Then _
            layoutCode = Mid(layoutCode, 1, tagStart - 6) & contentEndToDoCode & _
                Mid(layoutCode, tagEnd)
    End Function

End Class
%>