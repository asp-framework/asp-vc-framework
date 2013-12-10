htmlFilter
==========
`htmlFilter` &mdash; HTML 过滤

说明
----
>     string htmlFilter(string htmlString)
> HTML 过滤。

参数
----
> `htmlString`
>> **类型：** `string`  
>> **说明：**HTML 标记字符串。

返回值
------
> **类型：**`string`  
> **说明：**纯文本。

范例
----
>
    <%
    Dim htmlString, resultString
    htmlString = _
        "<div style="">S</div>" & _
        "<span id=""se"">E</span>" & _
        "<img />"
>
    resultString = SE.module(String).htmlFilter(htmlString)
    Response.Write(resultString)
    %>
> 以上内容输出：
>
    SE