## Installation

```shell
$ composer require "rekkles/easyexcel" -vvv
```

## Usage
```php
use EasyExcel\Read\ExcelToArray;
use EasyExcel\Read\ChunkReadFilter;

//简单获取Excel的数据为Array
$getData = new ExcelToArray('/Users/rekkles/code/data.csv');
$getData->load();
var_dump($getData->getData());

//分批获取Excel的数据（防止内存泄漏）
$chunk = new ChunkReadFilter();
$chunk->setRows(10, 2);
$data = new ExcelToArray('/Users/rekkles/code/data.csv');
var_dump($data->loadByChunk($chunk)->getData());

//写入Excel 生成文件到指定目录
$f = new ArrayToExcel(array(
            'fileName' => 'test.csv',           //导出的excel的文件的名称
            'sheetTitle' => '11',         //每个工作薄的标题
            'creator' => 'rekkles',            //创建者
            'writeType' => 'CSV',           //输出类型 Excel5 Excel7 CSV
            'path' => ROOT_PATH.'runtime/',           //输出路径 确保有写入权限
));
$f->setFirstRow(array('',1111,2222,3333))
  ->fillData(array(['','aaa','bbb','ccc'],['','ddd','eee','fff']));
```

## Documentation
   目前是第一版，功能还在逐渐完善中。
   
   简单的一些功能在ExcelToArray中已经封装好函数，写好了中文文档。
   
## Feature
   1.批量导入、导出
   
## Donate  

   <img src="https://rekkles.xyz/wx.png" width="200" alt="EasyWeChat Logo"/>
   <img src="https://rekkles.xyz/alipay.jpeg" height="264" alt="EasyWeChat Logo"/>
   
  [个人博客](https://rekkles.xyz/)  
  [个人gitbook](https://gitbook.rekkles.xyz/)  
    
       

