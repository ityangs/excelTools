# excelTools
PHPExcel导入导出工具

## 优点

- 可以自由的导出Excel表到浏览器或者文件夹中
- 可以传入数组或者json数据，自动处理
- 方便命名导出的文件名
- 方便导出文件的后缀
- 可以按照自定义的模板导出Excel表格





## 使用方式
```php
header("content-type:text/html;charset=utf-8");
require_once 'api/Excel.php';
/* 封装数据（一般从数据库查询，json数据或者数组都可以） */
$headerRow=['a1','b1','c1','d1'];
$dataRow=[['a','b','c','d'],['a','b','c','d'],['a','b','c','d'],['a','b','c','d'],['a','b','c','d'],['a','b','c','d']];

echo Excel::export($headerRow, $dataRow, $exportName = "文件名称1", $exportFileType = "xlsx", $sheetTitle = "sheet1", $exportType = 2, $template = 'template.xls');
```
