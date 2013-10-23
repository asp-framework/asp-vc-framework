<%
'''
 ' SimpleExtensionsBase.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.10.23
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsBase

    ' @var dictionary configs <配置项>
    ' 获取函数: getConfigs
    Private configs

    ' @var boolean aspIncludeTag <是否开启ASP #include 标签>
    ' 设置函数: setAspIncludeTag
    ' 判断函数: isAspIncludeTag
    Private aspIncludeTag

    ' @var dictionary modulesQueue <模块队列>
    ' 设置函数: addModule
    ' 获取函数: getModule
    Private modulesQueue

'###########################'
'###########################'

    '''
     ' 构造函数
     ''
    Private Sub Class_Initialize

    End Sub

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
                Response.Write("[FUNCTION] LoadFile Error - 找不到檔案：" & filePath)
                Response.End
            End If
            loadFile = .ReadText
            .Close
        End With
        Set oStream = Nothing
    End Function

    '''
     ' 包含并运行文件
     '
     ' @param string filePath <文件路径>
     '
     ' @return string <可执行代码>
     ''
    Public Function include(ByVal filePath)
        Dim ASP_TAG_LEFT, ASP_TAG_RIGHT
        ASP_TAG_LEFT = "<" & "%" : ASP_TAG_RIGHT = "%" & ">"

        ' code: 存放包含文件转译后的可运行代码
        ' codeCache: 代码处理时的临时缓存
        ' content: 文件内容
        ' codeEnd: 标签内容结束位置
        ' codeStart: 标签内容开始位置
        Dim code, codeCache, content, codeEnd, codeStart

        ' 包含模式
        If isAspIncludeTag Then
            content = aspIncludeTagProcess(filePath)
        Else
            content = Me.loadFile(filePath)
        End If

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
                Case "@"
                    codeCache = Null
                Case "="
                    codeCache = "Response.Write(" & Mid(codeCache, 2) & ")"
            End Select

            code = code & codeCache & vbCrLf : codeCache = Null
            codeStart = InStr(codeEnd, content, ASP_TAG_LEFT) + 2
        Loop

        include = code : simpleExtensionsIncludeCodeExecute(code)
    End Function

    '''
     ' 设置ASP #include 标签是否开启
     '
     ' @param boolean isAspincludeTag <标签是否开启>
     ''
    Public Property Let setAspIncludeTag(ByVal isAspincludeTag)
        If VarType(isAspincludeTag) = 11 Then aspIncludeTag = isAspincludeTag
    End Property

    '''
     ' 获取ASP #include 标签是否开启
     '
     ' @return boolean <标签是否开启>
     ''
    Public Property Get isAspIncludeTag()
        If IsEmpty(aspIncludeTag) Then aspIncludeTag = True
        isAspIncludeTag = aspIncludeTag
    End Property

    '''
     ' ASP #include 的实现
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
        content = Me.loadFile(filePath)
        Do While True
            codeStart = InStr(1, content, ASP_INCLUDE_TAG_LEFT) + 4
            codeEnd = InStr(codeStart, content, ASP_INCLUDE_TAG_RIGHT) + 3

            ' 跳出解析
            If codeStart = 4 Then Exit Do

            contentCache = Trim(Mid(content, codeStart, codeEnd - codeStart - 3))
            If InStr(1,contentCache, "#include", 1) = 1 Then
                contentCache = Trim(Mid(contentCache, 9))
                filePath = Replace(filePath, "\", "/")
                If InStr(1,contentCache, "file", 1) = 1 Then
                    filePath = Mid(filePath,1,InstrRev(filePath,"/")) & Replace(Trim(Mid(Trim(Mid(contentCache, 5)), 2)), """", "")
                ElseIf InStr(contentCache, "virtual", 1) = 1 Then
                    filePath = Replace(Trim(Mid(Trim(Mid(contentCache, 8)), 2)), """", "")
                End If
                contentCache = Empty

                content = Mid(content, 1, codeStart - 5) & aspIncludeTagProcess(filePath) & Mid(content, codeEnd)
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
        configFilePath = Server.MapPath(configFilePath)

        Dim seConfigsDoc : Set seConfigsDoc = Server.CreateObject("Microsoft.XMLDOM")
        seConfigsDoc.Async = False
        seConfigsDoc.Load(configFilePath)
        Set seConfigsDoc = seConfigsDoc.getElementsByTagName("seConfigs")(0)
        Call processConfigs(seConfigsDoc, getConfigs(Null))
        Set seConfigsDoc = Nothing
    End Function

    '''
     ' 获取配置项
     ''
    Public Property Get getConfigs(ByVal configPath)
        If VarType(configs) <> 9 Then Set configs = Server.CreateObject("Scripting.Dictionary")

        If IsNull(configPath) Then
            Set getConfigs = configs
        End If
    End Property

    '''
     ' 处理载入的配置
     '
     ' @param object xmlDoc <XML数据>
     ' @param dictionary nowConfigs <配置项>
     ''
    Private Function processConfigs(ByRef xmlDoc, ByRef nowConfigs)
        If VarType(xmlDoc) <> 9 Then Exit Function

        Dim config, nowNode, attributes
        For Each nowNode In xmlDoc.childNodes
            Select Case nowNode.nodeType
                ' 元素
                Case 1
                    nowConfigs.Add nowNode.NodeName, Server.CreateObject("Scripting.Dictionary")

                    ' 节点属性
                    nowConfigs.Item(nowNode.NodeName).Add "SimpleExtensionsConfigAttributes", Server.CreateObject("Scripting.Dictionary")
                    For Each attributes In nowNode.Attributes
                        nowConfigs.Item(nowNode.NodeName).Item("SimpleExtensionsConfigAttributes").Add attributes.NodeName, Server.CreateObject("Scripting.Dictionary")
                        nowConfigs.Item(nowNode.NodeName).Item("SimpleExtensionsConfigAttributes").Item(attributes.NodeName) = attributes.NodeValue
                    Next

                    Call processConfigs(nowNode, nowConfigs.Item(nowNode.NodeName))
                ' 文本
                Case 3
                    nowConfigs.Add "SimpleExtensionsConfigText", nowNode.Text
            End Select
        Next
    End Function

    '''
     ' 调用模块
     '
     ' @param string moduleName <模块名称>
     ''
    Public Function module(ByVal moduleName)

    End Function

    '''
     ' 向队列增加模块
     '
     ' @param string moduleName <模块名称>
     ''
    Private Property Set addModule(ByVal moduleName)
        If VarType(modulesQueue) <> 9 Then Set modulesQueue = Server.CreateObject("Scripting.Dictionary")
        If Not modulesQueue.Exists(moduleName) Then
            Dim modulePath
            modulePath = getSEDir & "/" & moduleName & "/" & moduleName & ".asp"
            Me.include(modulePath)
            Set modulesQueue.Item(moduleName) = Eval("New " & moduleName)
        End If
    End Property

    '''
     ' 获取模块
     '
     ' @param string moduleName <模块名称>
     '
     ' @return class|Nothing <实例化的模块>
     ''
    Private Property Get getModule(ByVal moduleName)
        If modulesQueue.Exists(moduleName) Then Set getModule = modulesQueue.Item(moduleName)
    End Property

End Class
%>

<%
'''
 ' 执行代码
 '
 ' @param string code <可执行代码>
 ''
Function simpleExtensionsIncludeCodeExecute(ByRef simpleExtensionsIncludeCode)
    simpleExtensionsIncludeCode = "simpleExtensionsIncludeCode = Empty" & vbCrLf & simpleExtensionsIncludeCode
    Execute(simpleExtensionsIncludeCode)
End Function
%>