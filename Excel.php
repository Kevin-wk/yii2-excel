<?php
/**
 * @link http://www.lubanr.com/
 * @copyright Copyright (c) 2015 Baochen Tech. Co. 
 * @license http://www.lubanr.com/license/
 */

namespace lubaogui\excel;

use Yii;
use yii\base\Exception;

/**
 * Yii Excel扩展
 * @author Baogui Lu (lbaogui@lubanr.com)
 * @version since 2.0
 */

require_once __DIR__. '/classes/PHPExcel.php';

class Excel 
{
    private $objPHPExcel = null;

    public function __construct() {

        $this->objPHPExcel = new \PHPExcel();

    }

    /**
     * @brief 将数组导出成为excel文件记录
     *
     * @param array $data 数组,数组的第一个元素是数组的相关模板信息
     * @param string $type 导出的excel文件类型
     * @return  public function 
     * @retval   
     * @see 
     * @note 
     * @author 吕宝贵
     * @date 2016/01/02 23:46:49
    **/
    public function exportToExcel(array $data, array $meta, $type = 1) {

        //字符集转换，excel使用gbk字符集
        $callbackFunc = function(&$value, $key) {
            mb_convert_encoding($value, 'gbk', 'utf-8');
        };

        array_walk($meta, 'callbackFunc');

        $this->objPHPExcel->getProperties()->setCreator($meta['author'])
            ->setLastModifiedBy($meta['modify_user'])
            ->setTitle($meta['title'])
            ->setSubject($meta['subject'])
            ->setDescription($meta['description'])
            ->setKeywords($meta['keywords'])
            ->setCategory($meta['category']);

        $this->objPHPExcel->setActiveSheetIndex(0);
        $objActiveSheet = $this->objPHPExcel->getActiveSheet();
        $objActiveSheet->setTitle($meta['title']);

        $columnCount = count($data[0]);

        $rowIndex = 0;
        foreach ($data as $payable) {
            //Excel的第A列，uid是你查出数组的键值，下面以此类推
            $startCharAscii = 65;
            for ($columnIndex = 0; $columnIndex < $columnCount; $columnIndex++) {
                $currentCharAscii = $startCharAscii + 1;
                $objActiveSheet->setCellValue(chr($currentCharAscii) . $rowIndex, $payable[$columnIndex]);
            }
            $rowIndex += 1;
        }

        $filename = $meta['filename'];
        $this->objPHPExcel->getActiveSheet()->setTitle('明细记录');
        $this->objPHPExcel->setActiveSheetIndex(0);
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="'.$filename.'.xls"');
        header('Cache-Control: max-age=0');
        $objWriter = \PHPExcel_IOFactory::createWriter($this->objPHPExcel, 'Excel5');

        $objWriter->save('php://output');

    }

    /**
     * @brief 从excel文件中导入记录，返回数组
     *
     * @param FILE|ContentString $file  具体的文件内容，必须为excel文件
     * @return array 从文件导出的数组信息 
     * @see 
     * @note 
     * @author 吕宝贵
     * @date 2016/01/02 23:46:14
    **/
    public function load($file) {

    }

}
