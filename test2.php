<?php

	
	date_default_timezone_set('Asia/Shanghai');
	include 'PHPExcel/PHPExcel.php';
    include 'PHPExcel/PHPExcel/Reader/Excel2007.php';
    include 'PHPExcel/PHPExcel/Writer/Excel5.php';
    include 'PHPExcel/PHPExcel/Writer/Excel2007.php';
    $filename = 'aaa.xlsx';//需要读取数据的excel表
    $file_type = 'xlsx';
    //读取excel
    if ($file_type == 'xlsx') {
        $objReader = PHPExcel_IOFactory::createReader('Excel2007');
    } else {
        $objReader = PHPExcel_IOFactory::createReader('Excel5');
    }
    $objReader->setReadDataOnly(true);
    $objPHPExcel= $objReader->load($filename);
    $objWorksheet= $objPHPExcel->getSheet();
    $hightestrow= $objWorksheet->getHighestRow();//总行数
    $highestColumn= $objWorksheet->getHighestColumn();//总列数
    $highestColumnIndex= PHPExcel_Cell::columnIndexFromString($highestColumn);
    $excelData= array();
    //获取excel数据
    for($row = 2; $row <= $hightestrow;$row++){
        for($col=0;$col<=$highestColumnIndex;$col++){
            $excelData[$row][] = (string)$objWorksheet->getCellByColumnAndRow($col,$row)->getValue();
        }
    }
    //如有需要，可链接数据库重新拼装数据
    foreach ($excelData as $key => $value) {
        $arr[$key][0] = $value[0];
        $arr[$key][1] = $value[1];
        $arr[$key][2] = $value[2];
        $arr[$key][3] = $value[3];
        $arr[$key][4] = $value[4];
        $arr[$key][5] = $value[5];
        $arr[$key][6] = 'Shanghai';
     }
     //将数组写入excel
    $objPHPExcels = new PHPExcel();
    $objSheets = $objPHPExcels->getActiveSheet(); 
    $objSheets->setTitle('helen');
    //  $j=2;
    //  foreach($arr as $val){  
    //     $objSheets->setCellValue('A'.$j,$val['0'])->setCellValue('B'.$j,$val['1'])->setCellValue('C'.$j,$val['2'])->setCellValue('D'.$j,$val['3'])->setCellValue('E'.$j,$val['4'])->setCellValue('F'.$j,$val['5'])->setCellValue('G'.$j,$val['6']);  
    //     $j++; // 每循环一次换一行写入数据  
    // }  
    // print_r("<pre>");
    // print_r($arr);
    // print_r("</pre>");die;
 	$objSheets->fromArray($arr);//数组方式添加
 	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcels,'Excel2007');
 	$objWriter->save('welldone.xlsx');
    