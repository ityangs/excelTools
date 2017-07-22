<?php
header("content-type:text/html;charset=utf-8");
require_once 'api/Excel.php';
/* 封装数据（一般从数据库查询，json数据或者数组都可以） */
$headerRow=['a1','b1','c1','d1'];
$dataRow=[['a','b','c','d'],['a','b','c','d'],['a','b','c','d'],['a','b','c','d'],['a','b','c','d'],['a','b','c','d']];

echo Excel::export($headerRow, $dataRow, $exportName = "文件名称1", $exportFileType = "xlsx", $sheetTitle = "sheet1", $exportType = 2, $template = 'template.xls');
































/* 模拟远程生成Excel表格 */
/* 客户端curl模拟提交代码 */

/**
 * 
 * @param unknown $url 远程URL接口地址
 * @param array $data 
 *   【* @param array headerRow 表格首行标题数据            
      * @param array dataRow 数据行数据             
      * @param type exportName 导出文件名称           
      * @param type exportFileType 导出格式  
      * @param type sheetTitle  工作表名称          
      * @param type exportType   导出类型,1:导出到浏览器 2:导出到服务器文件系统（assets文件夹下，后期将文件生成到七牛上）                  
      * @param type template 模板文件名(例如：xxx.xls或者xxx.xlsx)】
 * @param string $json
 * @return multitype:boolean number
 */
/* function http($url, $data=NULL, $json = false)
{
    $curl = curl_init();
    curl_setopt($curl, CURLOPT_URL, $url);
    curl_setopt($curl, CURLOPT_SSL_VERIFYPEER, false);
    curl_setopt($curl, CURLOPT_SSL_VERIFYHOST, false);
    if (!empty($data)) {
        curl_setopt($curl, CURLOPT_POST, 1);
        curl_setopt($curl, CURLOPT_POSTFIELDS, http_build_query($data));
        if ($json && is_array($data)) { // 发送JSON数据
            $data = json_encode($data);
            curl_setopt($curl, CURLOPT_HEADER, 0);
            curl_setopt($curl, CURLOPT_HTTPHEADER, array(
                'Content-Type: application/json; charset=utf-8',
                'Content-Length:' . strlen($data)
            ));
        }
    }
    curl_setopt($curl, CURLOPT_RETURNTRANSFER, 1);
    $res = curl_exec($curl);
    var_dump($res);
    $errorno = curl_errno($curl);
    
    if ($errorno) {
        return array(
            'errorno' => false,
            'errmsg' => $errorno
        );
    }
    curl_close($curl);

    return json_decode($res, true); 
} */








