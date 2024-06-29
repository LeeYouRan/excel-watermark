<?php

namespace Services;

use ZipArchive;
use Exception;

/**
 * Class Watermark
 * 用于在基于OpenXML格式的Excel文件（xlsx）中添加水印
 *
 * @version 1.0
 *
 * @example
 *          $watermark = new \Services\Watermark('D:/a.xlsx');
 *          $imageNumber = $watermark->addImage('D:/images/b.png');
 *          $watermark->getSheet(1)->setBgImg($imageNumber);
 *          $watermark->close();
 *
 * @package Services
 */
class Watermark {

    /**
     * @var ZipArchive 用于处理zip文件的ZipArchive对象
     */
    private $zip;

    /**
     * @var int 图像序号，用于给每个添加的图片分配唯一编号
     */
    private $num = 1;

    /**
     * @var int sheet序号，用于指定当前操作的工作表
     */
    private $sheet = 1;

    /**
     * @var array 存储图像后缀名的数组
     */
    private $suffixArr = [];

    /**
     * @var array 存储图像唯一名称的数组
     */
    private $nameArr = [];

    /**
     * @var string 关系文件路径的格式字符串
     */
    private const RELS_PATH = 'xl/worksheets/_rels/sheet%s.xml.rels';

    /**
     * @var string 工作表文件路径的格式字符串
     */
    private const SHEET_PATH = 'xl/worksheets/sheet%s.xml';

    /**
     * @var string 媒体文件路径的格式字符串
     */
    private const MEDIA_PATH = 'xl/media/bgimage%s.%s';

    /**
     * 初始化
     *
     * @param string|null $file 压缩包文件名
     *
     * @throws Exception 如果文件无法打开，抛出异常
     */
    public function __construct(?string $file = null) {
        if (!empty($file)) {
            $this->openFile($file); // 如果提供了文件名，尝试打开文件
        }
    }

    /**
     * 打开zip文件
     *
     * @param string $file 压缩包文件名
     * @throws Exception 如果文件无法打开，抛出异常
     */
    private function openFile(string $file): void {
        $this->zip = new ZipArchive(); // 创建一个新的ZipArchive对象
        if ($this->zip->open($file) !== true) {
            throw new Exception("Unable to open the file: $file"); // 如果打开文件失败，抛出异常
        }
    }

    /**
     * 设置zip文件
     *
     * @param string $file 压缩包文件名
     * @return bool 成功打开文件返回true
     *
     * @throws Exception 如果文件无法打开，抛出异常
     */
    public function setFile(string $file): bool {
        $this->openFile($file); // 调用openFile方法打开文件
        return true; // 返回true表示成功
    }

    /**
     * 添加图片到zip文件中
     *
     * @param string $file 图片文件名
     * @return int 添加进来的图像的编号
     *
     * @throws Exception 如果无法添加图片，抛出异常
     */
    public function addImage(string $file): int {
        $suffix = pathinfo($file, PATHINFO_EXTENSION); // 获取图片文件的后缀名
        $name = uniqid(); // 生成一个唯一的图片名称
        if (!$this->zip->addFile($file, sprintf(self::MEDIA_PATH, $name, $suffix))) {
            throw new Exception("Unable to add image: $file"); // 如果无法添加图片，抛出异常
        }
        $num = $this->num; // 获取当前图片编号
        $this->suffixArr[$num] = $suffix; // 存储图片的后缀名
        $this->nameArr[$num] = $name; // 存储图片的唯一名称
        $this->num++; // 增加图片编号
        return $num; // 返回图片编号
    }

    /**
     * 获取指定的工作表
     *
     * @param int $num sheet编号
     * @return $this 返回当前对象实例
     */
    public function getSheet(int $num = 1): self {
        $this->sheet = $num; // 设置当前操作的工作表编号
        return $this; // 返回当前对象实例，支持方法链调用
    }

    /**
     * 设置背景图
     *
     * @param int $num 图像编号
     *
     * @throws Exception 如果图像编号无效，抛出异常
     */
    public function setBgImg(int $num): void {
        if (!isset($this->suffixArr[$num]) || !isset($this->nameArr[$num])) {
            throw new Exception("Invalid image number: $num"); // 如果图像编号无效，抛出异常
        }

        $nowSuffix = strtolower($this->suffixArr[$num]); // 获取图像后缀名并转为小写
        $relContent = $this->generateRelContent($num, $nowSuffix); // 生成关系文件的内容
        $this->addToZip(sprintf(self::RELS_PATH, $this->sheet), $relContent); // 将关系文件内容添加到zip文件中

        $sheetContent = $this->getUpdatedSheetContent(); // 获取更新后的工作表内容
        $this->addToZip(sprintf(self::SHEET_PATH, $this->sheet), $sheetContent); // 将更新后的工作表内容添加到zip文件中

        $contentTypes = $this->updateContentTypes($nowSuffix); // 更新内容类型文件
        $this->addToZip("[Content_Types].xml", $contentTypes); // 将更新后的内容类型文件添加到zip文件中
    }

    /**
     * 关闭zip文件
     */
    public function close(): void {
        if ($this->zip) {
            $this->zip->close(); // 关闭zip文件
        }
    }

    /**
     * 生成关系文件的内容
     *
     * @param int $num 图像编号
     * @param string $suffix 图像后缀名
     * @return string 生成的关系文件内容
     */
    private function generateRelContent(int $num, string $suffix): string {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/bgimage' . $this->nameArr[$num] . '.' . $suffix . '"/>
</Relationships>'; // 返回生成的关系文件内容
    }

    /**
     * 获取更新后的工作表内容
     *
     * @return string 更新后的工作表内容
     *
     * @throws Exception 如果无法获取工作表内容，抛出异常
     */
    private function getUpdatedSheetContent(): string {
        $sheetContent = $this->zip->getFromName(sprintf(self::SHEET_PATH, $this->sheet)); // 获取当前工作表的内容
        if ($sheetContent === false) {
            throw new Exception("Unable to get sheet content"); // 如果无法获取工作表内容，抛出异常
        }
        return str_replace('</worksheet>', '<picture r:id="rId1"/></worksheet>', $sheetContent); // 在工作表内容中添加图片标记并返回更新后的内容
    }

    /**
     * 更新内容类型文件
     *
     * @param string $suffix 图像后缀名
     * @return string 更新后的内容类型文件内容
     *
     * @throws Exception 如果无法获取内容类型文件，抛出异常
     */
    private function updateContentTypes(string $suffix): string {
        $contentTypes = $this->zip->getFromName('[Content_Types].xml'); // 获取内容类型文件的内容
        if ($contentTypes === false) {
            throw new Exception("Unable to get content types"); // 如果无法获取内容类型文件，抛出异常
        }
        if (!strpos($contentTypes, 'Extension="' . $suffix . '"') || !strpos($contentTypes, 'ContentType="image/' . $suffix . '"')) {
            $contentTypes = str_replace('</Types>', '<Default ContentType="image/' . $suffix . '" Extension="' . $suffix . '"/></Types>', $contentTypes); // 如果内容类型文件中不包含当前图像的后缀名，添加相应的内容类型
        }
        return $contentTypes; // 返回更新后的内容类型文件内容
    }

    /**
     * 添加内容到zip文件
     *
     * @param string $path 文件路径
     * @param string $content 文件内容
     *
     * @throws Exception 如果无法添加内容到zip文件，抛出异常
     */
    private function addToZip(string $path, string $content): void {
        if (!$this->zip->addFromString($path, $content)) {
            throw new Exception("Unable to add content to zip: $path"); // 如果无法添加内容到zip文件，抛出异常
        }
    }
}
