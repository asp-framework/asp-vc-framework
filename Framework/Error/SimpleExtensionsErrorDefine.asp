<%
'''
 ' SimpleExtensionsErrorDefine.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.11.7
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsErrorDefine

    ' @var array <错误定义数组>
    Private errorDefine(3)

'###########################'
'###########################'

    Private Sub Class_Initialize
        errorDefine(0) = "系统正常"
        errorDefine(1) = "用户自定义错误"
        errorDefine(2) = "系统错误"
    End Sub

    '''
     ' 获取错误编号的定义
     '
     ' @return string <错误编号的定义>
     ''
    Public Property Get getErrorDefine(ByVal errorDefineNumber)
        getErrorDefine = errorDefine(errorDefineNumber)
    End Property

End Class
%>