<?php
 	date_default_timezone_set('Asia/Shanghai');
	set_time_limit(0);  
    // $starttime = microtime();
    // echo microtime();
    $filename = 'aaa.xlsx';

    include 'PHPExcel/PHPExcel.php';
    include 'PHPExcel/PHPExcel/Reader/Excel2007.php';
    include 'PHPExcel/PHPExcel/Writer/Excel5.php';
    include 'PHPExcel/PHPExcel/Writer/Excel2007.php';
    $file_type = 'xlsx';
    if ($file_type == 'xlsx') {
        $objReader = PHPExcel_IOFactory::createReader('Excel2007');
    } else {
        $objReader = PHPExcel_IOFactory::createReader('Excel5');
    }
    $objReader->setReadDataOnly(true);
    $objPHPExcel= $objReader->load($filename);
    $objWorksheet= $objPHPExcel->getSheet();
    $hightestrow= $objWorksheet->getHighestRow();
    $highestColumn= $objWorksheet->getHighestColumn();
    $highestColumnIndex= PHPExcel_Cell::columnIndexFromString($highestColumn);
    $excelData= array();
    for($row = 2; $row <= $hightestrow;$row++){
    	for($col=0;$col<=$highestColumnIndex;$col++){
    		$excelData[$row][] = (string)$objWorksheet->getCellByColumnAndRow($col,$row)->getValue();
    	}
    }
    
    $phoneArr = [];
    foreach ($excelData as $key => $value) {
    	$phoneArr[] = $value[5];
    }
 //    print_r("<pre>");
	// print_r($phoneArr);
	// print_r("</pre>");
    
    // foreach ($phoneArr as $key => $value) {
    // $value = 13641283851;
    // 	$sms = array('province'=>'', 'supplier'=>''); 
    // 	$url = "http://tcc.taobao.com/cc/json/mobile_tel_segment.htm?tel=".$value."&t=".time();
    // 	$content = file_get_contents($url);
    // 	$sms['province'] = substr($content, "56", "4");
    // 	$sms['supplier'] = substr($content, "81", "4");
    	
    // }
    // $endtime = microtime();
    // $thistime = $endtime-$starttime;
    // $thistime = round($thistime,3);
    // echo $thistime;
    
    // for($i=0;$i<=intval(count($phoneArr)/10*1);$i++){

    // }
    $link= @mysqli_connect('localhost','root','000000') or die("mysql failed!");
    mysqli_select_db($link,'sina_picture');
    mysqli_set_charset($link,'utf8');
    
    foreach ($excelData as $key =>$value ){
    	$mts = substr($value[5],0,7);
		$sql = "select province,city from phonezone where mts=$mts";
		$result = mysqli_query($link,$sql);
		$row = mysqli_fetch_assoc($result);
		$arr[$key][0] = $value[0];
        $arr[$key][1] = $value[1];
        $arr[$key][2] = $value[2];
        $arr[$key][3] = $value[3];
        $arr[$key][4] = $value[4];
        $arr[$key][5] = $value[5];
        $arr[$key][6] = $row['province'].$row['city'];
	}
    
	$objPHPExcels = new PHPExcel();
    $objSheets = $objPHPExcels->getActiveSheet(); 	
    $objSheets->setTitle('helen');
    $j=2;
     foreach($arr as $val){  
        $objSheets->setCellValue('A'.$j,$val['0'])->setCellValue('B'.$j,$val['1'])->setCellValue('C'.$j, $val['2'])->setCellValue('D'.$j,$val['3'])->setCellValue('E'.$j,$val['4'])->setCellValue('F'.$j,$val['5'])->setCellValue('G'.$j,$val['6']);  
        $j++; // 每循环一次换一行写入数据  
    }  
 	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcels,'Excel2007');
 	$objWriter->save('zhengshi.xlsx');