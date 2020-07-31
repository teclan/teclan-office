# 文档工具

## FreeMarker

>用于输出内容到固定模板的Word文档

具体参考官网[https://freemarker.apache.org](https://freemarker.apache.org)
或中文网 [http://freemarker.foofun.cn](http://freemarker.foofun.cn)

- 操作步骤
 - 自定义一个模板，参考 `template/word/工作证明.docx`，其中 `${key}` 表示后续会使用`key`的真实内容替换此表达式
 - 将 `template/word/工作证明.docx`另存为 `template/word/工作证明.xml`
 - 将`template/word/工作证明.xml` 另存为或直接改名为 `template/word/工作证明.ftl`，`.ftl`即为程序需要的word模板  
 - 使用方法参考 `src/test/java/com/teclan/office/word/WordFactory.workProveTest()`

- 常见问题
 - 输出带表格的word以上转换可能存在异常，详情请参考[https://www.cnblogs.com/w-yu-chen/p/11402098.html](https://www.cnblogs.com/w-yu-chen/p/11402098.html)
 或`documnets/freemarker导出带表格Word文档异常处理.xps`
 
 ## OpenOffice
 
 > [http://www.openoffice.org/download/](http://www.openoffice.org/download/)
 
 - windows
 
 ``` 
cd C:\Program Files (x86)\OpenOffice 4\program
soffice -headless -accept="socket,host=127.0.0.1,port=8100;urp;
```