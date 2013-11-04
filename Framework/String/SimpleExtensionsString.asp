<%
'''
 ' SimpleExtensionsString.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.11.4
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<!-- 导入支持文件 -->
    <!-- #include file = "SimpleExtensionsStringMD5.asp" -->
<!-- /导入支持文件 -->

<%
Class SimpleExtensionsString

    ' @var class <MD5类>
    Private md5Class

    '''
     ' MD5加密
     '
     ' @param string stringToMD5 <需要处理成MD5的字符串>
     '
     ' @return string <MD5字符串>
     ''
    Public Function md5(ByVal stringToMD5)
        If VarType(md5Class) <> 9 Then Set md5Class = New SimpleExtensionsStringMD5
        md5 = md5Class.md5(stringToMD5)
    End Function

End Class
%>