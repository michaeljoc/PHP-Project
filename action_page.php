<?php

	
	require_once dirname(__FILE__) . '\PHPExcel-1.8\Classes\PHPExcel.php';

	require_once dirname(__FILE__) . '\PHPExcel-1.8\Classes\PHPExcel\IOFactory.php';
	
	$fn = $_POST["firstname"];
	$mn = $_POST["middlename"];
	$ln = $_POST["lastname"];
	$bd = $_POST["bday"];
	$sc = $_POST["school"];
	
	$cl = $_POST["class"];
	$ha = $_POST["address"];
	$pc = $_POST["postcode"];
	$mob = $_POST["mobile"];
	$ema = $_POST["email"];
	
	$gen = $_POST["gender"]; 
	$born = $_POST["borninaus"];
	$abt = $_POST["abnltor"];
	$eng = $_POST["english"];
	$lang = $_POST["language"];
	
	$sha = $_POST["shareresults"];
	$conc = $_POST["concerns"];
	$condet = $_POST["condetails"];
	$eainf = $_POST["earinfect"];
	$surg = $_POST["earsurgery"];
	
	$tes = $_POST["heartest"];
	$hld = $_POST["heardiag"];
	$condaf = $_POST["condiaffect"];
	$condafdet = $_POST["condiaffectdetails"];
	$contact = $_POST["agreement"];
	
	$parentname = $_POST["parentname"];
	$signature = $_POST["signature"];
	
	
	
	$objPHPExcel = new PHPExcel();
	
	$objPHPExcel = PHPExcel_IOFactory::load('secondtest.xlsx'); 
	

	$sheet = $objPHPExcel->getActiveSheet();
	$highestRow = $sheet->getHighestRow();
	
	$x = $highestRow + 1;

	$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A'.$x, $fn)
            ->setCellValue('B'.$x, $mn)
            ->setCellValue('C'.$x, $ln)
            ->setCellValue('D'.$x, $bd)
			->setCellValue('E'.$x, $sc)
			
			->setCellValue('F'.$x, $cl)
			->setCellValue('G'.$x, $ha)
			->setCellValue('H'.$x, $pc)
			->setCellValue('I'.$x, $mob)
			->setCellValue('J'.$x, $ema)
			
			->setCellValue('K'.$x, $gen)
			->setCellValue('L'.$x, $born)
			->setCellValue('M'.$x, $abt)
			->setCellValue('N'.$x, $eng)
			->setCellValue('O'.$x, $lang)
			
			->setCellValue('P'.$x, $sha)
			->setCellValue('Q'.$x, $conc)
			->setCellValue('R'.$x, $condet)
			->setCellValue('S'.$x, $eainf)
			->setCellValue('T'.$x, $surg)
			
			->setCellValue('U'.$x, $tes)
			->setCellValue('V'.$x, $hld)
			->setCellValue('W'.$x, $condaf)
			->setCellValue('X'.$x, $condafdet)
			->setCellValue('Y'.$x, $contact)
			
			->setCellValue('Z'.$x, $parentname)
			->setCellValue('AA'.$x, $signature);
//	}
	
	
	$sheet = $objPHPExcel->getActiveSheet();
	$highestRow = $sheet->getHighestRow();
	
	


	
	
	
	$callStartTime = microtime(true);

	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
	
	$objWriter->save('secondtest.xlsx');

	$callEndTime = microtime(true);
	$callTime = $callEndTime - $callStartTime;
	
		
?>