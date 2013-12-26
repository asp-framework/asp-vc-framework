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

    ' @var string <内容开始前要执行的代码>
    ' (视图文件 '<!-- #contentStart -->' 标签 前的代码)
    Private contentStartToDoCode

    ' @var string <内容结束后要执行的代码>
    ' (视图文件 '<!-- #contentEnd -->'标签 后的代码)
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
        parseSETag()

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
     ' 解析视图中的特殊标签
     ''
    Private Function parseSETag()
        parseContentStartTag()
        parseContentEndTag()
    End Function

    '''
     ' 解析 '<!-- #contentStart -->' 标签
     ' 此函数直接操作 contentCode, contentStartToDoCode 参数
     ''
    Private Function parseContentStartTag()
        Dim CONTENT_START_TAG_LEFT, CONTENT_START_TAG, CONTENT_START_TAG_RIGHT
        CONTENT_START_TAG_LEFT = "'<!--"
        CONTENT_START_TAG = "#contentStart"
        CONTENT_START_TAG_RIGHT = "-->'"

        Dim tagStart, tagEnd
        tagStart = InStr(1, contentCode, CONTENT_START_TAG_LEFT) + 5
        If tagStart = 5 Then Exit Function
        tagEnd = InStr(tagStart, contentCode, CONTENT_START_TAG_RIGHT) + 4

        Dim searchedContentEndTag
        searchedContentEndTag = Trim(Mid(contentCode, tagStart, tagEnd - tagStart - 4))
        If InStr(1, searchedContentEndTag, CONTENT_START_TAG, 1) = 1 Then
            contentStartToDoCode = Mid(contentCode, 1, tagStart-6)

            ' 清除结尾的换行
            If InStrRev(contentStartToDoCode, vbCrLf) = Len(contentStartToDoCode)-1 Then _
                contentStartToDoCode = Left(contentStartToDoCode, Len(contentStartToDoCode)-2)
            If InStrRev(contentStartToDoCode, " & vbCrLf)") = Len(contentStartToDoCode)-9 Then _
                contentStartToDoCode = Left(contentStartToDoCode, Len(contentStartToDoCode)-10) & ")"

            contentCode = Mid(contentCode, tagEnd)
        End If
    End Function

    '''
     ' 解析 '<!-- #contentEnd -->' 标签
     ' 此函数直接操作 contentCode, contentEndToDoCode 参数
     ''
    Private Function parseContentEndTag()
        Dim CONTENT_END_TAG_LEFT, CONTENT_END_TAG, CONTENT_END_TAG_RIGHT
        CONTENT_END_TAG_LEFT = "'<!--"
        CONTENT_END_TAG = "#contentEnd"
        CONTENT_END_TAG_RIGHT = "-->'"

        Dim tagStart, tagEnd
        tagStart = InStr(1, contentCode, CONTENT_END_TAG_LEFT) + 5
        If tagStart = 5 Then Exit Function
        tagEnd = InStr(tagStart, contentCode, CONTENT_END_TAG_RIGHT) + 4

        Dim searchedContentEndTag
        searchedContentEndTag = Trim(Mid(contentCode, tagStart, tagEnd - tagStart - 4))
        If InStr(1, searchedContentEndTag, CONTENT_END_TAG, 1) = 1 Then
            contentEndToDoCode = Mid(contentCode, tagEnd)

            ' 清除开头的换行
            If InStrRev(contentEndToDoCode, vbCrLf) = Len(contentEndToDoCode)-1 Then _
                contentEndToDoCode = Left(contentEndToDoCode, Len(contentEndToDoCode)-2)
            If InStrRev(contentEndToDoCode, "(vbCrLf & ") = 17 Then _
                contentEndToDoCode = "Response.Write(" & Mid(contentEndToDoCode, 27)

            contentCode = Mid(contentCode, 1, tagStart-6)
        End If
    End Function

    '''
     ' 替换布局中的特殊标签
     ''
    Private Function replaceSETag()
        replaceContentStartTag()
        replaceContentTag()
        replaceContentEndToDoTag()
    End Function

    '''
     ' 替换布局中的 '<!-- #contentStartToDo -->' 标签
     ''
    Private Function replaceContentStartTag()
        Dim CONTENT_START_TODO_TAG_LEFT, CONTENT_START_TODO_TAG, CONTENT_START_TODO_TAG_RIGHT
        CONTENT_START_TODO_TAG_LEFT = "'<!--"
        CONTENT_START_TODO_TAG = "#contentStartToDo"
        CONTENT_START_TODO_TAG_RIGHT = "-->'"

        Dim tagStart, tagEnd
        tagStart = InStr(1, layoutCode, CONTENT_START_TODO_TAG_LEFT) + 5
        If tagStart = 5 Then Exit Function
        tagEnd = InStr(tagStart, layoutCode, CONTENT_START_TODO_TAG_RIGHT) + 4

        Dim searchedContentendToDoTag
        searchedContentendToDoTag = Trim(Mid(layoutCode, tagStart, tagEnd - tagStart - 4))
        If InStr(1, searchedContentendToDoTag, CONTENT_START_TODO_TAG, 1) = 1 Then _
            layoutCode = Mid(layoutCode, 1, tagStart - 6) & contentStartToDoCode & _
                Mid(layoutCode, tagEnd)
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
        Dim CONTENT_END_TODO_TAG_LEFT, CONTENT_END_TODO_TAG, CONTENT_END_TODO_TAG_RIGHT
        CONTENT_END_TODO_TAG_LEFT = "'<!--"
        CONTENT_END_TODO_TAG = "#contentEndToDO"
        CONTENT_END_TODO_TAG_RIGHT = "-->'"

        Dim tagStart, tagEnd
        tagStart = InStr(1, layoutCode, CONTENT_END_TODO_TAG_LEFT) + 5
        If tagStart = 5 Then Exit Function
        tagEnd = InStr(tagStart, layoutCode, CONTENT_END_TODO_TAG_RIGHT) + 4

        Dim searchedContentendToDoTag
        searchedContentendToDoTag = Trim(Mid(layoutCode, tagStart, tagEnd - tagStart - 4))
        If InStr(1, searchedContentendToDoTag, CONTENT_END_TODO_TAG, 1) = 1 Then _
            layoutCode = Mid(layoutCode, 1, tagStart - 6) & contentEndToDoCode & _
                Mid(layoutCode, tagEnd)
    End Function

End Class
%>