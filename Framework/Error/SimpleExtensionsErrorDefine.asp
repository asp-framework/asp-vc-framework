<%
'''
 ' SimpleExtensionsErrorDefine.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.11.6
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsErrorDefine

    ' @var array <错误定义数组>
    Private errorDefine(0)

'###########################'
'###########################'

    Private Sub Class_Initialize
        errorDefine(0) = "无错误"
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