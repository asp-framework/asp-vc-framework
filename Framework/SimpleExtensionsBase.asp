<%
'''
 ' SimpleExtensionsBase.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.12.26
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsBase

    ' @var dictionary <配置项>
    ' 获取函数: getConfigs
    Private configs

    ' @var dictionary <模块队列>
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
                Call Me.module("Error").throwError( _
                    2, _
                    "导入文件【" & filePath & "】失败" _
                )
            End If
            On Error GoTo 0
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
     ' 包含文件并获取可执行代码(不执行内容)
     '
     ' @param string filePath <文件路径>
     '
     ' @return string <可执行代码>
     ''
    Public Function getIncludeCode(ByVal filePath)
        getIncludeCode = pressModeInclude(filePath, 2)
    End Function

    '''
     ' 包含文件并获取执行后的内容(不输出内容)
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
                    Dim fileName
                    fileName = Replace(Trim(Mid(Trim(Mid(contentCache, 5)), 2)), """", "")
                    filePath = Mid(filePath,1,InstrRev(filePath,"/")) & fileName
                ElseIf InStr(contentCache, "virtual", 1) = 1 Then
                    filePath = Replace(Trim(Mid(Trim(Mid(contentCache, 8)), 2)), """", "")
                End If
                contentCache = Empty

                ' 替换标签为文件内容
                content = Mid(content, 1, codeStart - 5) & _
                    aspIncludeTagProcess(filePath) & Mid(content, codeEnd)

                codeEnd = 1
            End If
        Loop

        aspIncludeTagProcess = content
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

        codeEnd = 1 : codeStart = InStr(codeEnd, content, ASP_TAG_LEFT)+2
        Do While True
            ' 输出非代码内容
            If codeStart = 2 Then
                codeCache = Mid(content, codeEnd)
            Else
                codeCache = Mid(content, codeEnd, codeStart-codeEnd-2)
            End If
            If Len(codeCache) Then
                codeCache = Replace(codeCache, """", """""")
                codeCache = Replace(codeCache, vbCrLf, """ & vbCrLf & """)
                codeCache = Replace(codeCache, vbLf, """ & vbCrLf & """)
                codeCache = "Response.Write(""" & codeCache & """)"
                codeCache = Replace(codeCache, "("""" & ", "(")
                codeCache = Replace(codeCache, "& """" &", "&")
                codeCache = Replace(codeCache, " & """")", ")")
                code = code & (codeCache & vbCrLf) : codeCache = Null
            End If

            ' 跳出解析
            If codeStart = 2 Then Exit Do

            codeEnd = InStr(codeStart, content, ASP_TAG_RIGHT)+2
            codeCache = Trim(Mid(content, codeStart, codeEnd-codeStart-2))

            ' 判断特殊标签
            Select Case Left(codeCache, 1)
                Case "@"
                    codeCache = Null
                Case "="
                    codeCache = Mid(codeCache, 2)
                    codeCache = "Response.Write(" & codeCache & ")"
            End Select

            code = code & (codeCache & vbCrLf) : codeCache = Null
            codeStart = InStr(codeEnd, content, ASP_TAG_LEFT)+2
        Loop
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

        If VarType(configs) <> 9 Then _
            Set configs = Server.CreateObject("Scripting.Dictionary")
        Call processConfigs(seConfigsDoc, configs)

        Set seConfigsDoc = Nothing

        checkConfigs()
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
            ' 元素节点
            If nowNode.nodeType = 1 Then
                If Not nowConfigs.Exists(nowNode.NodeName) Then _
                    Call nowConfigs.Add( _
                        nowNode.NodeName, _
                        Server.CreateObject("Scripting.Dictionary") _
                    )

                ' 节点属性
                If Not nowConfigs.Item(nowNode.NodeName).Exists("Attributes") Then _
                    Call nowConfigs.Item(nowNode.NodeName).Add( _
                        "Attributes", _
                        Server.CreateObject("Scripting.Dictionary") _
                    )
                For Each attributes In nowNode.Attributes
                    If nowConfigs.Item(nowNode.NodeName).Item("Attributes") _
                    .Exists(attributes.NodeName) Then
                        nowConfigs.Item(nowNode.NodeName).Item("Attributes") _
                            .Item(attributes.NodeName) = attributes.NodeValue
                    Else
                        Call nowConfigs.Item(nowNode.NodeName).Item("Attributes").Add( _
                            attributes.NodeName, attributes.NodeValue _
                        )
                    End If
                Next

                Call processConfigs(nowNode, nowConfigs.Item(nowNode.NodeName))

            ' 文本
            ElseIf nowNode.nodeType = 3 Then
                If nowConfigs.Exists("Value") Then
                    nowConfigs.Item("Value") = nowNode.Text
                Else
                    Call nowConfigs.Add("Value", nowNode.Text)
                End If
            End If
        Next
    End Function

    '''
     ' 检查基本配置是否设置
     ''
    Private Function checkConfigs()
        If IsEmpty(getConfigs("System/seDir/Value")) Then
            Response.Write("框架目录没有设置。")
            Response.End()
        End If
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
                If Len(nowPath) > 0 Then _
                    evalString = evalString & ".Item(""" & nowPath & """)"
            Next
            On Error Resume Next
            getConfigs = Eval(evalString)
            ' 配置项不存在的错误处理
            If Err.Number = 424 Or Len(getConfigs) = 0 Then
                getConfigs = Empty
                Err.Clear
            End If
            On Error GoTo 0
        End If
    End Property

    '''
     ' 获取框架根目录
     ''
    Public Property Get getSEDir()
        getSEDir = getConfigs("System/seDir/Value")
    End Property

    '''
     ' 判断是否开发环境
     '
     ' @return boolean <是否开发环境>
     ''
    Public Property Get isDevelopment()
        isDevelopment = getConfigs("System/development/Value")

        If IsEmpty(isDevelopment) Then
            isDevelopment = False
        ElseIf StrComp(isDevelopment, "True", 1) = 0 Then
            isDevelopment = True
        ElseIf StrComp(isDevelopment, "False", 1) = 0 Then
            isDevelopment = False
        Else
            isDevelopment = False
        End If
    End Property

    '''
     ' 调用模块
     '
     ' @param string moduleName <模块名称>
     '
     ' @return class|nothing <实例化的模块>
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
        If VarType(modulesQueue) <> 9 Then _
            Set modulesQueue = Server.CreateObject("Scripting.Dictionary")

        If modulesQueue.Exists(moduleName) Then Exit Function

        Dim modulePath
        modulePath = Me.getSEDir & "/" & moduleName & "/" & _
            "SimpleExtensions" & moduleName & ".asp"
        On Error Resume Next
        Response.Buffer = True
        Response.Flush()
        Me.include(modulePath)
        Response.Clear()
        ' 类重命名时的处理
        If Err.Number = 1041 Then On Error GoTo 0
        Call modulesQueue.Add( _
            moduleName, Eval("New " & "SimpleExtensions" & moduleName) _
        )
    End Function

    '''
     ' 获取模块
     '
     ' @param string moduleName <模块名称>
     '
     ' @return class <实例化的模块>
     ''
    Private Property Get getModule(ByVal moduleName)
        Set getModule = modulesQueue.Item(moduleName)
    End Property

End Class
%>