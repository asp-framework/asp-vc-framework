<%
'''
 ' SimpleExtensionsError.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.12.3
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<% ' 导入支持文件 %>
    <!-- #include file = "SimpleExtensionsErrorDefine.asp" -->
<% ' /导入支持文件 %>

<%
Class SimpleExtensionsError

    ' @var integer <异常错误编号>
    Private errorNumber

    ' @var class <异常错误定义类>
    Private errorDefineClass

    ' @var string <异常错误消息>
    Private errorMessage

    ' @var string <重定向URL>
    Private redirectURL

'###########################'
'###########################'

    Private Sub Class_Initialize
        ' 初始化异常错误定义类
        Set errorDefineClass = New SimpleExtensionsErrorDefine
        initConfigs()
    End Sub

    '''
     ' 初始化配置项
     ''
    Private Sub initConfigs()
        redirectURL = SE.getConfigs("Error/redirectURL/Value")
    End Sub

'###########################'
'###########################'

    '''
     ' 抛出异常错误
     '
     ' @param integer throwErrorNumber <异常错误编号>
     ' @param sting message <异常错误信息>
     ''
    Public Function throwError(ByVal throwErrorNumber, ByVal message)
        errorNumber = throwErrorNumber
        errorMessage = message
        If SE.isDevelopment Then
            Execute(SE.getIncludeCode(SE.getSEDir & "/" & "Error/Error.html"))
            Response.End()
        Else
            If Not IsEmpty(redirectURL) Then Response.Redirect(redirectURL)
        End If
    End Function

    '''
     ' 获取当前异常错误编号
     '
     ' @return integer <当前异常错误编号>
     ''
    Public Property Get getError()
        If IsEmpty(errorNumber) Then errorNumber = 0
        getError = errorNumber
    End Property

    '''
     ' 获取异常错误编号定义
     '
     ' @param integer <异常错误编号>
     '
     ' @return string <异常错误编号定义>
     ''
    Public Property Get getErrorDefine(ByVal errorNumber)
        getErrorDefine = errorDefineClass.getErrorDefine(errorNumber)
    End Property

    '''
     ' 获取异常错误消息
     '
     ' @return string <异常错误消息>
     ''
    Public Property Get getErrorMessage()
        getErrorMessage = errorMessage
    End Property

    '''
     ' 设置重定向URL
     '
     ' @param string urlString <URL字符串>
     ''
    Public Function setRedirectURL(ByVal urlString)
        redirectURL = urlString
    End Function

    '''
     ' 获取重定向URL
     '
     ' @return string|null <重定向URL字符串> 
     ''
    Public Property Get getRedirectURL()
        getRedirectURL = redirectURL
    End Property

End Class
%>