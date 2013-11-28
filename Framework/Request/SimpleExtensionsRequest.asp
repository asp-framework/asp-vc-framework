<%
'''
 ' SimpleExtensionsRequest.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.11.28
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsRequest

    ' @var string <主机>
    Private host

    ' @var string <路径>
    Private path

    ' @var string <HTTP请求信息字符串>
    Private queryString

'###########################'
'###########################'

    Private Sub Class_Initialize
        init()
    End Sub

    Private Sub init()
        host = Request.ServerVariables("HTTP_HOST")
        path = Request.ServerVariables("PATH_INFO")
        queryString = Request.ServerVariables("QUERY_STRING")
    End Sub

    '''
     ' 获取URL。
     '
     ' @param string urlType <获取的URL类型>
     ''
    Public Function getUrl(ByVal urlType)

    End Function

    '''
     ' 获取URL,并赋上参数。
     ''
    Public Function getUrlWith()

    End Function

End Class
%>