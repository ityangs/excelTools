<?php

/**
 * 导出Excel
 * @date: 2017年7月21日 下午4:42:09
 * 
 * @author : ityangs<ityangs@163.com>
 *        
 */
class ExportExcel
{

    /**
     * 根路径
     * @var string
     */
     private $appRoot;
    /**
     * 使用PHPExcel链接库的路径
     *
     * @var string
     */
    private $libraryPath = "/libraries/PHPExcel.class.php";

    /**
     * 模板文件路径
     *
     * @var string
     */
    private $templatesPath = "/templates/";
    /**
     * 文件生成路径
     *
     * @var string
     */
    private $assetsPath = "/assets/";

    /**
     * 表格首行标题数据
     *
     * @var array
     */
    private $headerRow;

    /**
     * 数据行数据
     *
     * @var array(array)
     */
    private $dataRow;

    /**
     * 工作表索引
     *
     * @var int
     */
    private $sheetIndex;

    /**
     * 起始行索引(例如:2)
     *
     * @var int
     */
    private $rowIndex;

    /**
     * 起始列索引(例如:"A")
     *
     * @var string
     */
    private $columnIndex;

    /**
     * 模板文件名(例如：xxx.xls或者xxx.xlsx)
     *
     * @var string
     */
    private $template;

    /**
     * 工作表名称
     *
     * @var string
     */
    private $sheetTitle;

    /**
     * 导出类型,1:导出到浏览器 2:导出到服务器文件系统（assets文件夹下）
     *
     * @var type
     */
    private $exportType;

    /**
     * 导出格式
     *
     * @var string
     */
    private $exportFileType;

    /**
     * 导出文件名称
     *
     * @var string
     */
    private $exportName;

    public function __construct($headerRow = NULL, $dataRow = NULL, $exportName = "导出文件", $exportFileType = "xls", $sheetTitle = "sheet", $exportType = 1, $template = NULL, $sheetIndex = 0, $rowIndex = NULL, $columnIndex = NULL)
    {
        ignore_user_abort(true); // 后台运行
        set_time_limit(0); // 取消脚本运行时间的超时上限
        ini_set('memory_limit','500M');
        $this->appRoot=dirname(dirname(str_replace('\\', '/', __FILE__)));
        $this->libraryPath = $this->appRoot . $this->libraryPath;
        $this->templatesPath = $this->appRoot . $this->templatesPath;
        $this->headerRow = $headerRow;
        $this->dataRow = $dataRow;
        $this->sheetIndex = $sheetIndex;
        $this->rowIndex = $rowIndex;
        $this->columnIndex = $columnIndex;
        $this->template = $template;
        $this->sheetTitle = $sheetTitle;
        $this->exportType = $exportType;
        $this->exportFileType = $exportFileType;
        $this->exportName = $exportName;
    }

    /**
     * 导出Excel文件
     * @date: 2017年7月21日 下午4:42:09
     *
     * @author : ityangs<ityangs@163.com>
     */
    public function createExcel()
    {
        require_once $this->libraryPath; // 引入Excel库
        
        $objPHPExcel = new \PHPExcel(); // 初始化对象
        if (! empty($this->template)) { // 引入模板文件
            $objPHPExcel = \PHPExcel_IOFactory::load($this->templatesPath . $this->template);
        }
        
        // 设置文件信息
        $objPHPExcel->getProperties()
            ->setCreator("Excel")
            ->setLastModifiedBy("Excel")
            ->setTitle("Excel")
            ->setSubject("Excel")
            ->setDescription("Excel.")
            ->setKeywords("Excel")
            ->setCategory("Excel");
        // 设置活动状态sheet
        $objPHPExcel->setActiveSheetIndex($this->sheetIndex);
        // 获取当前活动的表
        $objActSheet = $objPHPExcel->getActiveSheet();
        // 起始列索引与起始行索引赋初始值
        if (empty($this->columnIndex) || empty($this->rowIndex)) {
            // 对前50*50遍历是否存在值,找出最大的列索引和行索引
            $maxJ = 0;
            $maxI = 0;
            for ($i = 1; $i <= 50; $i ++) {
                for ($j = 1; $j <= 50; $j ++) {
                    $cellNameTemp = $this->getColumnIndex($i) . $j;
                    $objCell = $objActSheet->getCell($cellNameTemp);
                    if (! empty($objCell->getValue())) {
                        $maxI = $maxI < $i ? $i : $maxI;
                        $maxJ = $maxJ < $j ? $j : $maxJ;
                    }
                }
            }
            // 起始列索引
            $this->columnIndex = $this->getColumnIndex($maxI + 1);
            // 起始行索引
            $this->rowIndex = $maxJ + 1;
        }
        $this->columnIndex = "A";
        
        // 列索引int类型
        $intColumnIndexTemp = \PHPExcel_Cell::columnIndexFromString($this->columnIndex);
        
        // 标题行数据录入
        if (! empty($this->headerRow)) {
            for ($i = 0; $i < count($this->headerRow) && $i <= 26 * 26; $i ++) {
                $strTemp = $this->getColumnIndex($i + $intColumnIndexTemp) . $this->rowIndex;
                $objActSheet->setCellValue($strTemp, $this->headerRow[$i]);
            }
            $this->rowIndex ++;
        }
        
        // 内容行数据录入
        for ($i = 0; $i < count($this->dataRow); $i ++) {
            $j = 0;
            foreach ($this->dataRow[$i] as $key => $value) {
                $strTemp = $this->getColumnIndex($j + $intColumnIndexTemp) . ($i + $this->rowIndex);
                $objActSheet->setCellValue($strTemp, $value);
                if (++ $j > 26 * 26) {
                    break;
                }
            }
        }
        
        // 输出
        if ($this->exportType == 1) {
            // 输出到浏览器
            header('Content-Type: application/vnd.ms-excel');
            header('Content-Disposition: attachment;filename="' . "$this->exportName.$this->exportFileType" . '"');
            header('Cache-Control: max-age=0');
            header('Cache-Control: max-age=1');
            header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
            header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
            header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
            header('Pragma: public'); // HTTP/1.0
            $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
            $objWriter->save('php://output');
            unset($objWriter);
        } elseif ($this->exportType == 2) {
                // 输出到服务器文件系统
                $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
                $fileName=iconv("utf-8", "gb2312",$this->appRoot.$this->assetsPath.$this->exportName.'.'.$this->exportFileType);
                $objWriter->save(str_replace('.php', '.'.$this->exportFileType,$fileName));
                
                $fileName=$this->appRoot.$this->assetsPath.$this->exportName.'.'.$this->exportFileType;
                return str_replace('.php', '.'.$this->exportFileType,$fileName);
                unset($objWriter);
            }
    }

    /**
     * 获取列索引(例如:AB) 同\PHPExcel_Cell::stringFromColumnIndex($i);以及columnIndexFromString
     *
     * @param int $number
     *            从1开始
     * @return :int @date: 2017年7月21日 下午4:51:34
     * @author : ityangs<ityangs@163.com>
     */
    private function getColumnIndex($number)
    {
        // A~Z共计26个字母,最多允许26*26=676列
        if ($number >= 1 && $number <= 26) {
            return chr(64 + $number);
        } else 
            if ($number > 26 && $number <= 26 * 26) {
                return chr(64 + $number / 26) . chr(64 + $number % 26);
            }
        return;
    }
    /**
    * 释放内存对象
    * @date: 2017年7月22日 上午10:06:29
    * @author: ityangs<ityangs@163.com>
    */
    public function __destruct(){
        unset($this);
    }
    
    
}


?>