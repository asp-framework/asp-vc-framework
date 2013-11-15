<%
'''
 ' SimpleExtensionsFile.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.11.15
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsFile

    '''
     ' 判断目录是否存在
     '
     ' @param string directory <目录>
     '
     ' @return boolean <目录是否存在>
     ''
    Public Function dirExists(ByVal directory)
        dirExists = False
        On Error Resume Next
        directory = Server.MapPath(directory)
        If Err.Number = -2147467259 Then
            Err.Clear
            On Error GoTo 0
            Exit Function
        End If

        Dim fileSystem : Set fileSystem = Server.CreateObject("Scripting.FileSystemObject")
        If fileSystem.FolderExists(directory) Then dirExists = True
        Set fileSystem = Nothing
    End Function

    '''
     ' 判断文件是否存在
     '
     ' @param string filePath <文件路径>
     '
     ' @return boolean <文件是否存在>
     ''
    Public Function fileExists(ByVal filePath)
        fileExists = False
        On Error Resume Next
        filePath = Server.MapPath(filePath)
        If Err.Number = -2147467259 Then
            Err.Clear
            On Error GoTo 0
            Exit Function
        End If

        Dim fileSystem : Set fileSystem = Server.CreateObject("Scripting.FileSystemObject")
        If fileSystem.FileExists(filePath) Then fileExists = True
        Set fileSystem = Nothing
    End Function

End Class
%>