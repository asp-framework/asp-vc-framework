getSEDir
========
`getSEDir` &mdash; 获取框架根目录

说明
----
>     string getSEDir(void)
> 获取框架根目录。

参数
----
> 没有参数。

返回值
------
> **类型：**`string`  
> **说明：**框架根目录。

范例
----
>
    <!-- 配置文件 -->
    <?xml version="1.0" encoding="UTF-8"?>
    <SEConfigs>
        <System>
            <development>True</development>
            <seDir>../Framework</seDir>
            <appsDir>Apps</appsDir>
        </System>
        ...
    </SEConfigs>
>>
>
    <% 
    Dim seDir
    seDir = SE.getSEDir
    Response.Write(seDir)
    %>
> 以上内容输出：
>
    ../Framework