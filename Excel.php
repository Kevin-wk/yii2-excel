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

require_once './classes/PHPExcel.php';

class Excel 
{
    private $objPHPExcel = null;

    public function __construct() {

        $this->objPHPExcel = new \PHPExcel();

    }

    /**
     * @brief 将数组导出成为excel文件记录
     *
     * @param array $data 数组
     * @param string $type 导出的excel文件类型
     * @return  public function 
     * @retval   
     * @see 
     * @note 
     * @author 吕宝贵
     * @date 2016/01/02 23:46:49
    **/
    public function exportArrayToExcel(array $data, $type ) {
        $this->objPHPExcel->getProperties()->setCreator('Mr-Hug')
            ->setLastModifiedBy('Mr-Hug')
            ->setTitle('银行批量付款' . date(time()))
            ->setSubject('数银行批量付')
            ->setDescription('银行付款记录数据')
            ->setKeywords('excel')
            ->setCategory('result file');

        int $num = 1;
        foreach ($data as $payable) {
            $objPHPExcel->setActiveSheetIndex(0)
                //Excel的第A列，uid是你查出数组的键值，下面以此类推
                ->setCellValue('A'.$num, $payable['uid'])    
                ->setCellValue('B'.$num, $payable['email'])
                ->setCellValue('C'.$num, $payable['password'])
                ->setCellValue('D'.$num, $payable['password'])
                ->setCellValue('E'.$num, $payable['password']);
            $num += 1;
        }

        $filename = '待付款明细' . date('Y-m-d', time());
        $objPHPExcel->getActiveSheet()->setTitle('付款明细');
        $objPHPExcel->setActiveSheetIndex(0);
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="'.$filename.'.xls"');
        header('Cache-Control: max-age=0');
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
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
    public function importFromFile($file) {

    }

}
