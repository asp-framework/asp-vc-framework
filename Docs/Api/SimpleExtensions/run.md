run
===
`run` &mdash; 运行框架

说明  
----
>     void run(string configFilePath)
> 启动框架。(使用此方法必须设置配置文件变量)

参数
----
> `configFilePath`  
>> **类型：**`string`  
>> **说明：**配置文件路径。(路径格式为相对路径，相对路径起始于首个调用文件目录。)  
>> **范例：**`"Configs/config.xml"`

返回值
------
> 没有返回值。

范例
----
>     <% SE.run("Configs/config.xml") %>