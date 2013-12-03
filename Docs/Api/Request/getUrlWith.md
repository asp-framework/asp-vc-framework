getUrlWith
==========
`getUrlWith` &mdash; 获取URL，并赋上参数

说明
----
>     string|null getUrlWith(string|integer urlType, string|null commandQueryString)
> 获取URL，并赋上参数。

参数
----
> `urlType`
>> **类型：**`string` | `integer`  
>> **说明：**获取的URL类型。
>>> 目录：`"Dir"` 或 `0`  
>>> 路径：`"Path"` 或 `1`  
>>> 目录 + 询问参数：`"DirWith"` 或 `2`  
>>> 路径 + 询问参数：`"PathWith"` 或 `3`
>
> `commandQueryString`
>> **类型：**`string` | `null`  
>> **说明：**询问命令字符串。  
>>> 命令之间使用 `"&"` 为间隔符  
>>> 减少：`"-参数"`  
>>> 修改(增加)：`"参数=值"`
>>
>> **范例：**`"-k&a=b"`

返回值
------
> **类型：**`string` | `null`  
> **说明：**URL字符串。

范例
----
> 测试URL：[http://localhost/home/index.asp?a=1&b=2](#)
>
    <%
    Dim url1, url2
    url1 = SE.module("Request").getUrlWith("Dir", "c=b")
    url1 = SE.module("Request").getUrlWith(0, "c=b")
    Response.Write(url1 & "<br />" & url2 "<br />")
>
    url1 = SE.module("Request").getUrlWith("Path", "c=b")
    url1 = SE.module("Request").getUrlWith(1, "c=b")
    Response.Write(url1 & "<br />" & url2 "<br />")
>
    url1 = SE.module("Request").getUrlWith("DirWith", "-a&b=a&c=b")
    url1 = SE.module("Request").getUrlWith(2, "-a&b=a&c=b")
    Response.Write(url1 & "<br />" & url2 "<br />")
>
    url1 = SE.module("Request").getUrlWith("PathWith", "-a&b=a&c=b")
    url1 = SE.module("Request").getUrlWith(3, "-a&b=a&c=b")
    Response.Write(url1 & "<br />" & url2)
    %>
> 以上内容输出：
>
    /home/?c=b
    /home/?c=b
    /home/index.asp?c=b
    /home/index.asp?c=b
    /home/?b=a&c=b
    /home/?b=a&c=b
    /home/index.asp?b=a&c=b
    /home/index.asp?b=a&c=b