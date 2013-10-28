<%
'''
 ' SimpleExtensionsString.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.10.28
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<!-- 导入支持文件 -->
    <!-- #include file = "SimpleExtensionsStringMD5.asp" -->
<!-- /导入支持文件 -->

<%
Class SimpleExtensionsString

    Public Function md5(ByVal stringToMD5)
        Set md5Class = New SimpleExtensionsStringMD5
        md5 = md5Class.md5(stringToMD5)
    End Function

End Class
%>