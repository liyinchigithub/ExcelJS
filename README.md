# ExcelJS-demo

ExcelJS 是一个 Node.js 模块，可用来读写和操作 XLSX 和 JSON 电子表格数据和样式

# 开发环境

|名称|版本|
|-|-|
|nodejs|v12.13.0|
|request|2.88.2|
|axios|0.23|
|path|0.12.7|
|exceljs|1.9.1|
|node-xlsx|0.15.0|
|fs|0.0.1-security|

## 说明
SheetJS仅支持excel格式文件为xlsx、csv 无法支持xls，可借助node-xlsx来解决此问题！

## 安装

```shell
npm install
```

## mocha

```shell
mocha
```

## node

```shell
node ./ExecelJS/until/excel2json.js
```

```shell
node ./ExecelJS/until/json2excel.js
```
##  数据格式


```json
{
	"fileName":"./导入付款结果模板.xlsx",
	"sheetName":"数据列表",
	"data":{
		"A2":"这是要写入的A2",
		"B2":"这是要写入的B2",
		"C2":"这是要写入的C2"
	},
	"importUrl":""
}
```


## 常见问题

1.报错“(node:5587) UnhandledPromiseRejectionWarning: TypeError: Cannot read property 'getCell' of undefined”
原因：sheetName区分大小写，要与excel的实际sheet名称一致。

2.
原因：



# nodejs中几个excel模块的简单对比

exceljs (支持复杂导出，功能齐全；文档写的太烂，反正我是看了大半天，github地址)

ejsexcel (支持复杂导出，功能齐全；国内大牛的开源项目，基于ejs模板渲染，github地址)

node-xlsx (不支持复杂导出；基于js-xlsx，功能比较简单，github地址)

excel-export (不支持复杂导出；需要一个xml作为导出模板，比较麻烦；且超过10个月没维护，github地址)


# excel文件对象介绍
1. workbook 对象
指整个Excel 文件，使用插件读取 Exce文件后就会获得 workbook 对象。

2. worksheet 对象
指Excel 文件中的表（sheet），一个Excel文件中包含n张表（n >= 1），而每张表对应的就是 worksheet 对象。

3. clumn 对象
指表中的列（纵向），每一列就是一个 clumn 对象。

4. row 对象
指表中的行（横向），每一行就是一个 row 对象。

5. cell 对象
指表中的单元格，每一个单元格就是一个 cell 对象。


# 有道笔记API

>http://note.youdao.com/s/cRN3hpQ0

# 博客

>参考：https://www.jianshu.com/p/8aa148435499
>参考：https://www.jianshu.com/p/0f6a338c54f4
>参考：https://www.jianshu.com/p/09c338cdb7de