<?php

class Excel{
    public function __construct(){
        header("Content-Type:text/html;charset=UTF-8");
        $this->CI = &get_instance();

        $this->CI->load->library('PHPExcel/PHPExcel');
        $this->CI->load->library('PHPExcel/PHPExcel/IOFactory.php');
//        require_once 'PHPExcel/PHPExcel.php';
//        require_once 'PHPExcel/PHPExcel/IOFactory.php';
    }

    public function import($inputFile, $fileTyle = '',$check = TRUE){
        if($check){
            $types = ['xlsx', 'xls'];
            if (!in_array($fileTyle, $types)) {
                return false;
            }
        }

        $inputFileType = PHPExcel_IOFactory::identify($inputFile);
        $objReader = PHPExcel_IOFactory::createReader($inputFileType);
        $objPHPExcel = $objReader->load($inputFile);

        $sheetData = $objPHPExcel->getActiveSheet()->toArray(null, true, true, true);

        return $sheetData;
    }

    public function export($headBind, $datas, $filename, $config = array()){
        $phpExcel = new PHPExcel();
        $getSheet = $phpExcel->getSheet();

        if(isset($config['width'])){
            foreach ($config['width'] as $k=>$v){
                $getSheet->getColumnDimension( $v['column'])->setWidth($v['width']);
            }
        }

        $max = count($datas);
        $startIndex = 1;
        for ($i = 0; $i < $max; $i++) {
            $row = $datas[$i];
            $rowIndex = $i + $startIndex;
            foreach ($headBind as $k => $v) {
                $getSheet->setCellValue($k . $rowIndex, $row[$v] ?: "");
            }
        }

        header('Content-Type: application/vnd.ms-excel; charset=utf-8');
        header("Content-Disposition: attachment;filename=".urlencode($filename).".xlsx");
        header('Cache-Control: max-age=0');
        $objWriter = \PHPExcel_IOFactory::createWriter($phpExcel, 'Excel2007');
        $objWriter->save("php://output");
        exit;
    }
}