<%
'''
 ' SimpleExtensionsRequest.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.11.30
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

    ' @var dictionary <URL类型>
    Private urlTypes

'###########################'
'###########################'

    Private Sub Class_Initialize
        init()
    End Sub

    Private Sub init()
        host = Request.ServerVariables("HTTP_HOST")
        path = Request.ServerVariables("PATH_INFO")
        If Len(Request.ServerVariables("QUERY_STRING")) > 0 Then _
            queryString = "?" & Request.ServerVariables("QUERY_STRING")

        Set urlTypes = Server.CreateObject("Scripting.Dictionary")
        Call urlTypes.Add("Dir", 0)
        Call urlTypes.Add("Path", 1)
        Call urlTypes.Add("DirWith", 2)
        Call urlTypes.Add("PathWith", 3)
    End Sub

    '''
     ' 获取URL。
     '
     ' @param string|integer urlType <获取的URL类型>
     '
     ' @return string|null <URL字符串|空值>
     ''
    Public Function getUrl(ByVal urlType)
        getUrl = Null
        Dim urlTypeValue : urlTypeValue = getUrlTypeValue(urlType)

        Select Case urlTypeValue
            Case 0 : getUrl = Left(path, InStrRev(path, "/"))
            Case 1 : getUrl = path
            Case 2 : getUrl = Left(path, InStrRev(path, "/")) & queryString
            Case 3 : getUrl = path & queryString
        End Select
    End Function

    '''
     ' 获取URL,并赋上参数。
     '
     ' @param string|integer urlType <获取的URL类型>
     ' @param string|null commandQueryString <询问命令字符串>
     '
     ' @return string|null <URL字符串|空值>
     ''
    Public Function getUrlWith(ByVal urlType, ByVal commandQueryString)
        getUrlWith = Null
        Dim urlTypeValue : urlTypeValue = getUrlTypeValue(urlType)

        ' 目录式 + QueryString
        If urlTypeValue = 0 Or urlTypeValue = 2 Then _
            getUrlWith = Left(path, InStrRev(path, "/")) & _
                executeCommandQueryString(urlTypeValue, commandQueryString)

        ' 路径式 + QueryString
        If urlTypeValue = 1 Or urlTypeValue = 3 Then _
            getUrlWith = path & executeCommandQueryString(urlTypeValue, commandQueryString)
    End Function

    '''
     ' 获取URL类型值。
     '
     ' @param string|integer urlType <获取的URL类型>
     '
     ' @return integer|null <URL类型值|空值>
     ''
    Private Function getUrlTypeValue(ByVal urlType)
        getUrlTypeValue = Null
        If IsNumeric(urlType) Then
            getUrlTypeValue = urlType
        Else
            If Not urlTypes.Exists(urlType) Then _
                Exit Function
            getUrlTypeValue = urlTypes.Item(urlType)
        End If
    End Function

    '''
     ' 执行询问命令
     '
     ' @param integer urlTypeValue <获取的URL类型值>
     ' @param string|null commandQueryString <询问命令字符串>
     '
     ' @return string|null <执行命令后的 QueryString>
     ''
    Private Function executeCommandQueryString(ByVal urlTypeValue, ByVal commandQueryString)
        executeCommandQueryString = Null

        Dim cacheArray, cacheArrayValue, equalIndex
        cacheArray = Split(commandQueryString, "&")

        ' 不带 QueryString 的路径
        If urlTypeValue = 0 Or urlTypeValue = 1 Then
            For Each cacheArrayValue In cacheArray
                If StrComp(Left(cacheArrayValue, 1), "-") Then
                    If InStr(cacheArrayValue, "=") Then
                        executeCommandQueryString = _
                            executeCommandQueryString & _
                            "&" & cacheArrayValue
                    Else
                        executeCommandQueryString = _
                            executeCommandQueryString & _
                            "&" & cacheArrayValue & "=" & _
                            Request.QueryString(cacheArrayValue)
                    End If
                End If
            Next
        End If

        ' 带 QueryString 的路径
        If urlTypeValue = 2 Or urlTypeValue = 3 Then

        End If

        executeCommandQueryString = Replace( _
            executeCommandQueryString, _
            "&", _
            "?", _
            1, _
            1 _
        )
    End Function

End Class
%>