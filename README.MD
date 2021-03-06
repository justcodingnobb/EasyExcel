
# 已停止维护，建议使用 xlswriter [地址](https://xlswriter-docs.viest.me/)


## Installation

```shell
$ composer require "rekkles/easyexcel" -vvv
```

## Usage
```php
use EasyExcel\Read\ExcelToArray;
use EasyExcel\Read\ChunkReadFilter;

//简单获取Excel的数据为Array
$config = ['firstRowAsIndex' => true];
$getData = new ExcelToArray('/Users/rekkles/code/data.csv',$config);
$getData->load();
var_dump($getData->getData());

//分批获取Excel的数据（防止内存泄漏）
$chunk = new ChunkReadFilter();
$chunk->setRows(10, 2);
$data = new ExcelToArray('/Users/rekkles/code/data.csv');
var_dump($data->loadByChunk($chunk)->getData());

//写入Excel 生成文件到指定目录
$outObj = new ArrayToExcel(array(
     'fileName' => 'test.csv',           //导出的excel的文件的名称
     'sheetTitle' => '11',              //每个工作薄的标题
     'creator' => 'rekkles',            //创建者
     'writeType' => 'CSV',              //输出类型 Excel5 Excel7 CSV
     'path' => ROOT_PATH.'file/',    //输出路径 确保有写入权限
));
$outObj->setFirstRow(array('',1111,2222,3333))
  ->fillData(array(
  ['','aaa','bbb','ccc'],
  ['','ddd','eee','fff']
  ));
```

## Documentation
   目前是第一版，功能还在逐渐完善中。
   
   代码不是很复杂，针对PHPExcel封装了一部，
   遇到问题可以自行debug，有bug提交PR就更好了！   
## Feature
   1.批量导入、导出
   2.支持更多的文件格式
   
## Donate  

   <img src="https://rekkles.xyz/wx.png" width="200" alt="wechat" style="float:left"/>
   <img src="https://rekkles.xyz/alipay.jpeg" height="264" alt="alipay" style="padding-left:20px"/>
   
  [个人博客](https://rekkles.xyz/)  
  [个人gitbook](https://gitbook.rekkles.xyz/)  
    
       

