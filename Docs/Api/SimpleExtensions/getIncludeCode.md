getIncludeCode
==============
`getIncludeCode` &mdash; 包含文件并获取可执行代码(不执行内容)

说明
----
>     string getIncludeCode(ByVal filePath)
> 包含文件并获取可执行代码(不执行内容)

参数
----
> `filePath`
>> **类型：**`string`  
>> **说明：**文件路径。(路径格式为相对路径，相对路径起始于首个调用文件目录。)  
>> **范例：**`"dir/filePath"`

返回值
------
> **类型：**`string`  
> **说明：**可执行代码字符串。

范例
----
>
    目录文件
>
    Dir/
        file.asp
    index.asp
>>
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
    Dim codeString
    codeString = SE.getIncludeCode("Dir/file.asp")
    Execute(codeString)
    %>
> 运行`index.asp`将得到以下内容：  
>
    成功执行。