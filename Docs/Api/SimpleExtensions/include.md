include
=======
`include` &mdash; 包含并执行文件

说明
----
>     void include(string filePath)
> 包含并执行文件。

参数
----
> `filePath`
>> **类型：**`string`  
>> **说明：**文件路径。(路径格式为相对路径，相对路径起始于首个调用文件目录。)  
>> **范例：**`"dir/filePath"`

返回值
------
> 无返回值。

范例
----
>
    目录文件
>
    Dir/
        file.asp
    index.asp
> **范例一**：
>
    <%
    '''
     ' Dir/file.asp 文件
     ''
    Response.Write("成功执行。")
    %>
>>
>
    <%
    '''
     ' index.asp 文件
     ''
    SE.include("Dir/file.asp")
    %>
> 运行`index.asp`将得到以下内容：  
>
    成功执行。
> **范例二**：
>
    <%
    '''
     ' Dir/file.asp 文件
     ''
    SE.include("../index.asp")
    %>
>>
>
    <%
    '''
     ' index.asp 文件
     ''
    Response.Write("成功执行。")
    %>
> 运行`Dir/file.asp`将得到以下内容：  
>
    成功执行。