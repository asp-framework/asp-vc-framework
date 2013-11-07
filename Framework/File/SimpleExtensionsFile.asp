<%
'''
 ' SimpleExtensionsFile.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.11.7
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsFile

    '''
     ' 判断文件是否存在
     '
     ' @param string filePath <文件路径>
     '
     ' @return boolean <文件是否存在>
     ''
    Public Function fileExists(ByVal filePath)
        fileExists = False
        filePath = Server.MapPath(filePath)

        Dim fileSystem : Set fileSystem = Server.CreateObject("Scripting.FileSystemObject")
        If fileSystem.FileExists(filePath) Then fileExists = True
        Set fileSystem = Nothing
    End Function

End Class
%>