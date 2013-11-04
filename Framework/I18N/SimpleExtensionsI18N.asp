<%
'''
 ' SimpleExtensionsI18N.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.11.3
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsI18N

    ' @var string <国际化目录>
    Private i18nDir

    ' @var string <本地语言>
    Private localLanguage

    ' @var dictionary <翻译的内容>
    Private tContent

'###########################'
'###########################'

    '''
     ' 构造函数
     ''
    Private Sub Class_Initialize
        i18nDir = SE.module("Controller").getAppDir & "/I18N"
        setLocalLanguage(SE.getConfigs("I18N/language/Value"))
    End Sub

    '''
     '  翻译指定信息到当前语言
     ''
    Public Function t(ByVal keyPath)
        If IsNull(keyPath) Then
            Set t = tContent
        Else
            keyPath = Replace(keyPath, "\", "/")
            Dim pathArray, nowPath, evalString
            pathArray = Split(keyPath, "/")
            evalString = "tContent"
            For Each nowPath In pathArray
                If Len(nowPath) > 0 Then evalString = evalString & ".Item(""" & nowPath & """)"
            Next
            t = Eval(evalString)
        End If
    End Function

'###########################'
'###########################'

    '''
     '  设置当前语言
     ''
     Public Function setLocalLanguage(ByVal language)
        loadTContent(language)
        localLanguage = language
     End Function

    '''
     '  读取翻译内容
     ''
     Private Function loadTContent(ByVal language)
        Dim i18nDoc : Set i18nDoc = Server.CreateObject("Microsoft.XMLDOM")
        i18nDoc.Async = False
        i18nDoc.Load(Server.MapPath(i18nDir & "/" & language & ".xml"))
        Set i18nDoc = i18nDoc.GetElementsByTagName("SEI18N")(0)

        Set tContent = Server.CreateObject("Scripting.Dictionary")
        Call processTContent(i18nDoc, tContent)

        Set i18nDoc = Nothing
     End Function

    '''
     ' 处理载入的配置
     '
     ' @param object xmlDoc <XML数据>
     ' @param dictionary nowTContent <配置项>
     ''
    Private Function processTContent(ByRef xmlDoc, ByRef nowTContent)
        If VarType(xmlDoc) <> 9 Then Exit Function

        Dim nowNode, attributes
        For Each nowNode In xmlDoc.ChildNodes
            Select Case nowNode.nodeType
                ' 元素节点
                Case 1
                    Call nowTContent.Add(nowNode.NodeName, Server.CreateObject("Scripting.Dictionary"))

                    ' 节点属性
                    Call nowTContent.Item(nowNode.NodeName).Add("Attributes", Server.CreateObject("Scripting.Dictionary"))
                    For Each attributes In nowNode.Attributes
                        Call nowTContent.Item(nowNode.NodeName).Item("Attributes").Add(attributes.NodeName, attributes.NodeValue)
                    Next

                    Call processTContent(nowNode, nowTContent.Item(nowNode.NodeName))
                ' 文本
                Case 3
                    Call nowTContent.Add("Value", nowNode.Text)
            End Select
        Next
    End Function

    '''
     '  获取当前语言
     ''
    Public Property Get getLocalLanguage()
        getLocalLanguage = localLanguage
    End Property

End Class
%>