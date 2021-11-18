# excel2word
读取excel，将每行数据转存到一个word模板文件上

第一次执行会生成三个目录在excel2word下：

- input
- output
- template

input是放数据文件的地方

output是导出word文件的地方

template是放word模板文件的地方

整个代码流程是：复制一个模板文件并打开，然后读取excel数据将每个单元格写到word文件表格的某个位置。

代码并不通用，仅供参考