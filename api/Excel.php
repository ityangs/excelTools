<?php
/**
 * Excel接口文件
 * @date: 2017年7月22日 上午9:11:03
 * @author: ityangs<ityangs@163.com>
 */
class Excel
{
/*     private $errorNum = [];        					//错误号
    private $errorMess="";       					//错误报告消息
    private $originName;   	     					//源文件名
    private $allowtype = array('xls','xlsx'); 	   //设置限制文件的类型
    private $maxsize = 100000000;  		 */			//限制文件大小（字节）
    


    
    /**
     * 导出Excel
     *
     * @param array $headerRow 表格首行标题数据（最好为json）            
     * @param array $dataRow 数据行数据  （最好为json）            
     * @param type $exportName 导出文件名称           
     * @param type $exportFileType 导出格式  
     * @param type $sheetTitle  工作表名称          
     * @param type $exportType   导出类型,1:导出到浏览器 2:导出到服务器文件系统（assets文件夹下，后期将文件生成到七牛上）                  
     * @param type $template 模板文件名(例如：xxx.xls或者xxx.xlsx)           
     * @return type @date: 2017年7月21日 下午5:07:05
     * @author : ityangs<ityangs@163.com>
     */
    public static function export($headerRow, $dataRow, $exportName = "文件名称", $exportFileType = "xls", $sheetTitle = "sheet", $exportType = 1, $template = '')
    {  
        $result = [];
        // 判断格式
        $headerRow=is_string($headerRow)?json_decode($headerRow,true):$headerRow;
        $dataRow=is_string($dataRow)?json_decode($dataRow,true):$dataRow;
        if (! is_array($headerRow) || ! is_array($dataRow)) {
            $result = [
                'status' => 400,
                'msg' => '数据格式错误！'
            ];
            return $result;
        }
        // 创建一个Excel对象
        require_once 'ExportExcel.class.php';
        $objExcle = new ExportExcel($headerRow, $dataRow, $exportName, $exportFileType, $sheetTitle, $exportType, $template);
        
        // 判断传输发送
        if ($exportType == 1) {
            $objExcle->createExcel();
        } elseif($exportType == 2) {
                return $objExcle->createExcel();
            }
    }
    
    
    
    /**
     * 导出Excel接口
     *
     * @param array $headerRow 表格首行标题数据
     * @param array $dataRow 数据行数据
     * @param type $exportName 导出文件名称
     * @param type $exportFileType 导出格式
     * @param type $sheetTitle  工作表名称 'sheet1'
     * @param type $exportType   导出类型,1:导出到浏览器 2:导出到服务器文件系统（assets文件夹下，后期将文件生成到七牛上）
     * @param type $template 模板文件名(例如：xxx.xls或者xxx.xlsx)
     * @return type @date: 2017年7月21日 下午5:07:05
     * @author : ityangs<ityangs@163.com>
     */
   /*  public  function export()
    {
        $params=$_REQUEST?$_REQUEST:'';
        $params=!empty($params) && is_array($params)?$params:json_decode($params,true);//接收的数据处理
        //数据重组
        $headerRow=isset($params['headerRow'])&&is_array($params['headerRow'])?$params['headerRow']:$this->errorNum[]=-1;
        $dataRow=isset($params['dataRow'])&&is_array($params['dataRow'])?$params['dataRow']:$this->errorNum[]=-2;
        $exportName=isset($params['headerRow'])&&!empty($params['headerRow'])?$params['headerRow']:"文件名称";
        $exportFileType=isset($params['exportFileType'])&&!empty($params['exportFileType'])?$params['exportFileType']:"xls";
        $exportName=isset($params['sheetTitle'])&&!empty($params['sheetTitle'])?$params['sheetTitle']:"sheet";
        $exportType=isset($params['exportType'])&&!empty($params['exportType'])?$params['exportType']:1;
        $template=isset($params['template'])&&!empty($params['template'])?$params['template']:'';
    
        $result = [];
        if (count($this->errorNum)>0) {
            $result=[
                'status'=>400,
                'message'=>$this->getError()
            ];
            return json_encode($result);
        }
    
        // 创建一个Excel对象
        require_once 'ExportExcel.class.php';
        $objExcle = new ExportExcel($headerRow, $dataRow, $exportName = "文件名称", $exportFileType = "xls", $sheetTitle = "sheet", $exportType = 1, $template = '');
    
        // 判断传输发送
        if ($exportType == 1) {
            $objExcle->createExcel();
        } elseif($exportType == 2) {
            return $objExcle->createExcel();
        }
    } */
    
    
    /**
     * 设置出错信息
     * @return:return_type
     * @date: 2017年7月22日 上午10:52:18
     * @author: ityangs<ityangs@163.com>
     */
   /*  private function getError() {
        $str = "处理Excel<font color='red'>{$this->originName}</font>时出错 : ";
        if(count($this->errorNum)>0){
            foreach ($this->errorNum as $k=>$v){
                switch ($this->errorNum) {
                    case 4: $str .= "没有Excel文件被导入"; break;
                    case 3: $str .= "文件只有部分被导入"; break;
                    case 2: $str .= "导入文件的大小超过了HTML表单中MAX_FILE_SIZE选项指定的值"; break;
                    case 1: $str .= "导入的文件超过了php.ini中upload_max_filesize选项限制的值"; break;
                    case -1: $str .= "导出文件的表格首行标题数据不能为空"; break;
                    case -2: $str .= "导出文件的表格数据行不能为空"; break;
                    default: $str .= "未知错误";
                }
            }
        }
    
        return $str.'<br>';
    } */
    
    
    
    
    
    
    
    
    
    
    
    
    
    
}














?>