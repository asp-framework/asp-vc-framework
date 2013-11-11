getConfigs
==========
`getConfigs` &mdash; 获取配置项

说明
----
>     dictionary|string|empty getConfigs(ByVal configPath)
> 获取配置项。

参数
----
> `configPath`
>> **类型：**`string` | `null`  
>> **说明：**配置项路径。  
>> **范例：**`"system/seDir/Value"`

返回值
------
> **类型：**`dictionary` | `string` | `empty`  
> **说明：**  
>> 当`configPath`参数为`null`时,返回`dictionary`;  
>> 当`configPath`参数为 **正确** 的配置项路径时,返回`string`;  
>> 当`configPath`参数为 **错误** 的配置项路径时,返回`empty`。

范例
----
>
    <!-- 配置文件 -->
    <?xml version="1.0" encoding="UTF-8"?>
    <SEConfigs>
        <System>
            <development>True</development>
            <seDir attributes="Attributes">../Framework</seDir>
            <appsDir>Apps</appsDir>
        </System>
    </SEConfigs>
>>
>
    <%
    Dim result
    ' 参数为null
    Set result = SE.getConfigs(Null)
    Response.Write(result.Item("System").Item("seDir").Item("Value"))
    Response.Write("<br />")
    Response.Write(result.Item("System").Item("seDir").Item("Attributes").Item("attributes"))
    Set result = Nothing
    %>
>
    <br />
>
    <%
    ' 参数为正确的配置路径
    result = SE.getConfigs("System/seDir/Value")
    Response.Write(result)
    Response.Write("<br />")
    result = SE.getConfigs("System/seDir/Attributes/attributes")
    Response.Write(result)
    %>
>
    <br />
>
    <%
    ' 参数为错误的配置路径
    result = SE.getConfigs("a")
    Response.Write(TypeName(result)
    %>
>   以上内容输出:
>
        ../Framework
        Attributes
        ../Framework
        Attributes
        Empty