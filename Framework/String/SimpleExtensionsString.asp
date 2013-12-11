<%
'''
 ' SimpleExtensionsString.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.12.10
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<% ' 导入支持文件 %>
    <!-- #include file = "SimpleExtensionsStringMD5.asp" -->
<% ' /导入支持文件 %>

<%
Class SimpleExtensionsString

    ' @var class <MD5类>
    Private md5Class

    '''
     ' MD5 加密
     '
     ' @param string stringToMD5 <需要处理成MD5的字符串>
     '
     ' @return string <MD5字符串>
     ''
    Public Function md5(ByVal stringToMD5)
        If VarType(md5Class) <> 9 Then Set md5Class = New SimpleExtensionsStringMD5
        md5 = md5Class.md5(stringToMD5)
    End Function

    '''
     ' HTML 过滤
     '
     ' @param string htmlString <HTML 标记字符串>
     '
     ' @return string <纯文本>
     ''
    Public Function htmlFilter(ByVal htmlString)
        Dim resultString
        Dim HTML_LEFT_TAG, HTML_RIGHT_TAG
        HTML_LEFT_TAG = "<" : HTML_RIGHT_TAG = ">"

        Dim tagStartPos, tagEndPos
        tagEndPos = 0 : tagStartPos = InStr(tagEndPos+1, htmlString, HTML_LEFT_TAG)
        Do While True
            If tagStartPos = 0 Then 
                resultString = resultString & Mid(htmlString, tagEndPos+1)
                Exit Do
            End If

            If tagStartPos-tagEndPos-1 > 0 Then _
                resultString = resultString & Mid(htmlString, tagEndPos+1, tagStartPos-tagEndPos-1)

            tagEndPos = Instr(tagStartPos+1, htmlString, HTML_RIGHT_TAG)
            tagStartPos = 0
            If tagEndPos > 0 Then _
                tagStartPos = InStr(tagEndPos+1, htmlString, HTML_LEFT_TAG)
            If tagStartPos = 0 Then Exit Do
        Loop

        htmlFilter = resultString
    End Function

End Class
%>