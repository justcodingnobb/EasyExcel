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
```

## Documentation
   目前是第一版，功能还在逐渐完善中。
   
   简单的一些功能在ExcelToArray中已经封装好函数，写好了中文文档。
   
## Feature
   1.导出Excel
   2.批量导入、导出
    
       

