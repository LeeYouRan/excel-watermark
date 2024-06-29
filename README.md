# excel-watermark
add watermark to excel use php php给excel添加背景图片水印


```php
// 创建Watermark实例并打开Excel文件
$watermark = new \Services\Watermark('D:/a.xlsx');

// 添加图片到Excel文件中，并获取图像编号
$imageNumber = $watermark->addImage('D:/images/b.png');

// 获取指定的工作表并设置背景图
$watermark->getSheet(1)->setBgImg($imageNumber);

// 关闭Excel文件
$watermark->close();
```
