<%
'''
 ' SimpleExtensionsError.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.11.6
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<!-- 导入支持文件 -->
    <!-- #include file = "SimpleExtensionsErrorDefine.asp" -->
<!-- /导入支持文件 -->

<%
Class SimpleExtensionsError

    ' @var integer <错误编号>
    Private errorNumber

    ' @var class <错误定义类>
    Private errorDefineClass

'###########################'
'###########################'

    Private Sub Class_Initialize
        ' 初始化错误定义类
        Set errorDefineClass = New SimpleExtensionsErrorDefine
    End Sub

'###########################'
'###########################'

    '''
     ' 抛出错误异常
     '
     ''
    Public Function throwError(ByVal throwErrorNumber, ByVal message)
        errorNumber = throwErrorNumber
        Execute(SE.getIncludeCode(SE.getSEDir & "/" & "Error/Error.html"))
    End Function

    '''
     ' 获取当前错误编号
     '
     ' @return integer <当前错误编号>
     ''
    Public Property Get getError()
        If IsEmpty(errorNumber) Then errorNumber = 0
        getError = errorNumber
    End Property

    '''
     ' 获取错误编号定义
     '
     ' @param integer <错误编号>
     '
     ' @return string <错误编号定义>
     ''
    Public Property Get getErrorDefine(ByVal errorNumber)
        getErrorDefine = errorDefineClass.getErrorDefine(errorNumber)
    End Property

End Class
%>