<%
'''
 ' SimpleExtensionsRequest.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.12.1
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
            queryString = Request.ServerVariables("QUERY_STRING")

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
            Case 2 : getUrl = Left(path, InStrRev(path, "/")) & "?" & queryString
            Case 3 : getUrl = path & "?" & queryString
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
        If urlTypeValue = 0 Or urlTypeValue = 1 Or IsEmpty(queryString) Then
            For Each cacheArrayValue In cacheArray
                executeCommandQueryString = _
                    executeCommandQueryString & _
                    noQueryStringValueProcess(cacheArrayValue)
            Next

        ' 带 QueryString 的路径
        ElseIf (urlTypeValue = 2 Or urlTypeValue = 3) Then
            executeCommandQueryString = "&" & queryString
            For Each cacheArrayValue In cacheArray
                executeCommandQueryString = hasQueryStringValueProcess( _
                    executeCommandQueryString, _
                    cacheArrayValue _
                )
            Next
        End If

        executeCommandQueryString = Replace( _
            executeCommandQueryString, _
            "&", _
            "?", _
            1, _
            1 _
        )
    End Function


    '''
     ' 不带 QueryString 的路径 参数处理
     '
     ' @param string value <需要处理的值>
     '
     ' @return string <处理后的QueryString项>
     ''
    Private Function noQueryStringValueProcess(ByVal value)
        If InStr(value, "-") = 1 Then Exit Function

        Dim queryString
        If InStr(value, "=") Then
            queryString = "&" & value

        ElseIf Len(Request.QueryString(value)) > 0 Then _
            queryString = _
                "&" & value & "=" & _
                Request.QueryString(value)
        End If

        noQueryStringValueProcess = queryString
    End Function

    '''
     ' 带 QueryString 的路径 参数处理
     '
     ' @param string queryString <处理的询问字符串>
     ' @param string value <需要处理的值>
     '
     ' @return string <处理后的询问字符串>
     ''
    Private Function hasQueryStringValueProcess(ByVal queryString, Byval value)
        Dim startPos, endPos

        startPos = InStr(queryString, "&" & value)
        If startPos Then _
            endPos = InStr(startPos+1, queryString, "&")-1

        If InStr(value, "-") = 1 Then
            startPos = InStr(queryString, "&" & Mid(value, 2))
            If startPos Then
                endPos = InStr(startPos+1, queryString, "&")-1
                If endPos > 0 Then
                    queryString = _
                        Left(queryString, startPos-1) & _
                        Mid(queryString, endPos+1)
                Else
                    queryString = Left(queryString, startPos-1)
                End If
            End If
        ElseIf InStr(value, "=") Then

        Else

        End If

        hasQueryStringValueProcess = queryString
    End Function

End Class
%>