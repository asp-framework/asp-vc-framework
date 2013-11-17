t
=
`t` &mdash; 翻译指定信息到当前语言

说明
----
>     dictionary|string t(string|null keyPath)
> 翻译指定信息到当前语言。

参数
----
> `keyPath`
>> **类型：**`string` | `null`  
>> **说明：**翻译项路径。  
>> **范例：**`"Body/content/Value"`。

返回值
------
> **类型：**`dictionary` | `string`  
> **说明：**翻译的内容数据|当前语言字符串。
>> 当传入的参数为`null`时，返回`dictionary`；  
>> 当传入的参数为`string`时，返回`string`。

范例
----
>
    <!-- 配置文件 -->
    <?xml version="1.0" encoding="UTF-8"?>
    <SEConfigs>
        ...
        <I18N>
            <language>zh-cn</language>
        </I18N>
        ...
    </SEConfigs>
>>
>
    <!-- 国际化翻译文件 -->
    <?xml version="1.0" encoding="UTF-8"?>
    <SEI18N>
        <Head>
            <title>SE</title>
        </Head>
        <Body>
            <content>你好～！</content>
        </Body>
    </SEI18N>
>>
>
    <%
    Dim title
    title = SE.module("I18N").t(Null).Item("Head").Item("title").Item("Value")
    Response.Write(title)
>
    Response.Write("<br />")
>
    Dim content
    content = SE.module("I18N").t("Body/content/Value")
    Response.Write(content)
    %>
> 以上内容输出：
>
    SE
    你好～！