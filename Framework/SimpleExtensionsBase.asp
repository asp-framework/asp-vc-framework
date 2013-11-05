<%
'''
 ' SimpleExtensionsBase.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.10.30
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsBase

    ' @var dictionary configs <配置项>
    ' 获取函数: getConfigs
    Private configs

    ' @var dictionary modulesQueue <模块队列>
    ' 设置函数: addModule
    ' 获取函数: getModule
    Private modulesQueue

'###########################'
'###########################'

    '''
     ' 读取文件
     '
     ' @param string filePath <文件路径>
     '
     ' @return string <文件内容>
     ''
    Public Function loadFile(ByVal filePath)
        Dim oStream
        On Error Resume Next
        Set oStream = Server.CreateObject("ADODB.Stream")
        With oStream
            .Type = 2
            .Mode = 3
            .CharSet = Response.Charset
            .Open
            .LoadFromFile(Server.MapPath(filePath))
            If Err.Number <> 0 Then
                Err.Clear
                Response.Write("[FUNCTION] loadFile Error - 找不到檔案：" & filePath)
                Response.End
            End If
            loadFile = .ReadText
            .Close
        End With
        Set oStream = Nothing
    End Function

    '''
     ' 包含并执行文件
     '
     ' @param string filePath <文件路径>
     ''
    Public Function include(ByVal filePath)
        Call pressModeInclude(filePath, 1)
    End Function

    '''
     ' 包含文件获取可执行代码,但不执行
     '
     ' @param string filePath <文件路径>
     '
     ' @return string <可执行代码>
     ''
    Public Function getIncludeCode(ByVal filePath)
        getIncludeCode = pressModeInclude(filePath, 2)
    End Function

    '''
     ' 包含文件获取执行后的内容,但不输出内容
     '
     ' @param string filePath <文件路径>
     '
     ' @return string <执行后的内容>
     ''
    Public Function getIncludeResult(ByVal filePath)
        getIncludeResult = pressModeInclude(filePath, 3)
    End Function

    '''
     ' 按模式包含
     '
     ' @param string filePath <文件路径>
     ' @param int mode <模式,
     '     1:包含并执行;
     '     2:包含文件获取可执行代码,但不执行;
     '     3:包含文件获取执行后的内容,但不输出内容;
     ' >
     ''
    Private Function pressModeInclude(ByRef filePath, ByVal mode)
        Dim code, result, content

        content = aspIncludeTagProcess(filePath)

        ' 处理包含的内容
        Call processIncludeContent(code, result, content)

        Select Case mode
            Case 1
                ExecuteGlobal(code)
            Case 2
                pressModeInclude = code
            Case 3
                code = Replace(code, "Response.Write", "result=result&", 1, -1, 0)
                Execute(code) : pressModeInclude = result
        End Select
    End Function

    '''
     ' 处理包含的内容
     '
     ' @param code string <存放包含文件转译后的可运行代码>
     ' @param result string <存放包含文件执行后的内容>
     ' @param content string <文件内容>
     ''
    Private Function processIncludeContent(ByRef code, ByRef result, ByRef content)
        Dim ASP_TAG_LEFT, ASP_TAG_RIGHT
        ASP_TAG_LEFT = "<" & "%" : ASP_TAG_RIGHT = "%" & ">"

        ' codeCache: 代码处理时的临时缓存
        ' codeEnd: 标签内容结束位置
        ' codeStart: 标签内容开始位置
        Dim codeCache, codeEnd, codeStart

        codeEnd = 1 : codeStart = InStr(codeEnd, content, ASP_TAG_LEFT) + 2
        Do While True
            ' 输出非代码内容
            If codeStart = 2 Then
                codeCache = Mid(content, codeEnd)
            Else
                codeCache = Mid(content, codeEnd, codeStart - codeEnd - 2)
            End If
            codeCache = Replace(codeCache, vbCrLf, """ & vbCrLf & """)
            codeCache = "Response.Write(""" & codeCache & """)"
            code = code & codeCache & vbCrLf : codeCache = Null

            ' 跳出解析
            If codeStart = 2 Then Exit Do

            codeEnd = InStr(codeStart, content, ASP_TAG_RIGHT) + 2
            codeCache = Trim(Mid(content, codeStart, codeEnd - codeStart - 2))

            ' 判断特殊标签
            Select Case Left(codeCache, 1)
                Case "@" : codeCache = Null
                Case "=" : codeCache = "Response.Write(" & Mid(codeCache, 2) & ")"
            End Select

            code = code & codeCache & vbCrLf : codeCache = Null
            codeStart = InStr(codeEnd, content, ASP_TAG_LEFT) + 2
        Loop
    End Function

    '''
     ' ASP #include 标签处理
     '
     ' @param string filePath <文件路径>
     '
     ' @return string <#include包含的所有内容>
     ''
    Private Function aspIncludeTagProcess(ByVal filePath)
        Dim ASP_INCLUDE_TAG_LEFT, ASP_INCLUDE_TAG_RIGHT
        ASP_INCLUDE_TAG_LEFT = "<!--" : ASP_INCLUDE_TAG_RIGHT = "-->"

        ' content: 文件内容
        ' contentCache: 文件内容处理时的临时缓存
        ' codeEnd: 标签内容结束位置
        ' codeStart: 标签内容开始位置
        Dim content, contentCache, codeEnd, codeStart
        codeEnd = 1
        content = Me.loadFile(filePath)
        Do While True
            codeStart = InStr(codeEnd, content, ASP_INCLUDE_TAG_LEFT) + 4
            codeEnd = InStr(codeStart, content, ASP_INCLUDE_TAG_RIGHT) + 3

            ' 跳出解析
            If codeStart = 4 Then Exit Do

            contentCache = Trim(Mid(content, codeStart, codeEnd - codeStart - 3))
            If InStr(1, contentCache, "#include", 1) = 1 Then
                contentCache = Trim(Mid(contentCache, 9))
                filePath = Replace(filePath, "\", "/")
                If InStr(1,contentCache, "file", 1) = 1 Then
                    filePath = Mid(filePath,1,InstrRev(filePath,"/")) & Replace(Trim(Mid(Trim(Mid(contentCache, 5)), 2)), """", "")
                ElseIf InStr(contentCache, "virtual", 1) = 1 Then
                    filePath = Replace(Trim(Mid(Trim(Mid(contentCache, 8)), 2)), """", "")
                End If
                contentCache = Empty

                ' 替换标签为文件内容
                content = Mid(content, 1, codeStart - 5) & aspIncludeTagProcess(filePath) & Mid(content, codeEnd)

                codeEnd = 1
            End If
        Loop

        aspIncludeTagProcess = content
    End Function

    '''
     ' 载入配置文件
     '
     ' @param string filePath <配置文件路径>
     ''
    Public Function loadConfigs(ByVal configFilePath)
        Dim seConfigsDoc : Set seConfigsDoc = Server.CreateObject("Microsoft.XMLDOM")
        seConfigsDoc.Async = False
        seConfigsDoc.Load(Server.MapPath(configFilePath))
        Set seConfigsDoc = seConfigsDoc.GetElementsByTagName("SEConfigs")(0)

        Set configs = Server.CreateObject("Scripting.Dictionary")
        Call processConfigs(seConfigsDoc, configs)

        Set seConfigsDoc = Nothing
    End Function

    '''
     ' 处理载入的配置
     '
     ' @param object xmlDoc <XML数据>
     ' @param dictionary nowConfigs <配置项>
     ''
    Private Function processConfigs(ByRef xmlDoc, ByRef nowConfigs)
        If VarType(xmlDoc) <> 9 Then Exit Function

        Dim nowNode, attributes
        For Each nowNode In xmlDoc.ChildNodes
            Select Case nowNode.nodeType
                ' 元素节点
                Case 1
                    Call nowConfigs.Add(nowNode.NodeName, Server.CreateObject("Scripting.Dictionary"))

                    ' 节点属性
                    Call nowConfigs.Item(nowNode.NodeName).Add("Attributes", Server.CreateObject("Scripting.Dictionary"))
                    For Each attributes In nowNode.Attributes
                        Call nowConfigs.Item(nowNode.NodeName).Item("Attributes").Add(attributes.NodeName, attributes.NodeValue)
                    Next

                    Call processConfigs(nowNode, nowConfigs.Item(nowNode.NodeName))
                ' 文本
                Case 3
                    Call nowConfigs.Add("Value", nowNode.Text)
            End Select
        Next
    End Function

    '''
     ' 获取配置项
     '
     ' @param null|string configPath <配置路径,例:"system/seDir/Value">
     '
     ' @return dictionary|string|empty <所有配置数据|配置项字符串|无数据>
     ''
    Public Property Get getConfigs(ByVal configPath)
        If IsNull(configPath) Then
            Set getConfigs = configs
        Else
            configPath = Replace(configPath, "\", "/")
            Dim pathArray, nowPath, evalString
            pathArray = Split(configPath, "/")
            evalString = "configs"
            For Each nowPath In pathArray
                If Len(nowPath) > 0 Then evalString = evalString & ".Item(""" & nowPath & """)"
            Next
            On Error Resume Next
            getConfigs = Eval(evalString)
            If Err.Number = 424 Then Err.Clear : getConfigs = Empty
        End If
    End Property

    '''
     ' 获取框架根目录
     ''
    Public Property Get getSEDir()
        getSEDir = getConfigs(Null).Item("System").Item("seDir").Item("Value")
    End Property

    '''
     ' 调用模块
     '
     ' @param string moduleName <模块名称>
     ''
    Public Function module(ByVal moduleName)
        addModule(moduleName)
        Set module = getModule(moduleName)
    End Function

    '''
     ' 向队列增加模块
     '
     ' @param string moduleName <模块名称>
     ''
    Private Function addModule(ByVal moduleName)
        If VarType(modulesQueue) <> 9 Then Set modulesQueue = Server.CreateObject("Scripting.Dictionary")

        If modulesQueue.Exists(moduleName) Then Exit Function

        Dim modulePath
        modulePath = getSEDir & "/" & moduleName & "/" & "SimpleExtensions" & moduleName & ".asp"
        On Error Resume Next
        Me.include(modulePath)
        ' 类重命名时的处理
        If Err.Number = 1041 Then Err.Clear
        Call modulesQueue.Add(moduleName, Eval("New " & "SimpleExtensions" & moduleName))

    End Function

    '''
     ' 获取模块
     '
     ' @param string moduleName <模块名称>
     '
     ' @return class|nothing <实例化的模块>
     ''
    Private Property Get getModule(ByVal moduleName)
        Set getModule = Nothing
        If modulesQueue.Exists(moduleName) Then Set getModule = modulesQueue.Item(moduleName)
    End Property

End Class
%>