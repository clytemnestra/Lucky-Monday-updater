<?php

use Application\TokenExtractor;

require_once "../vendor/autoload.php";


function getCellColor(PHPExcel_Worksheet $worksheet, $cell)
{
    return $worksheet->getStyle($cell)->getFill()->getStartColor()->getRGB();
}

function setCellColor(PHPExcel_Worksheet $worksheet, $cell, $color)
{
    return $worksheet->getStyle($cell)->getFill()->getStartColor()->setRGB($color);
}


function addLmColumn($lmNumber){

}

function getLmColumn($lmNumber){

}





$a = new TokenExtractor('responses.xlsx','#177');
$res = $a->getResultsArray();
echo "<pre>";
var_dump($res);
echo "</pre>";
echo "<pre>";
ksort($res,SORT_NUMERIC );
var_dump($res);
echo "</pre>";
exit;


//
//$file = 'sheet.xlsx';
//$sheetNameToLoad = 'Tokens';
//$headerRowNumber = 1;
//$currentTokensColumn= 'C';
//$nicknamesColumn = 'E';
//
//
///** @var PHPExcel_Reader_Excel2007 $reader */
//$reader = new PHPExcel_Reader_Excel2007();
//$reader->setLoadSheetsOnly($sheetNameToLoad);
//$objPHPExcel = $reader->load($file);
//$worksheet = $objPHPExcel->getActiveSheet();
//
//
//$highestRow = $worksheet->getHighestRow();
//$highestColumn = $worksheet->getHighestColumn();
//
//
//
//$myRow = 1;
//
//$a= $worksheet
//    ->rangeToArray(
//        'A' . $myRow .
//        ':' .
//        $worksheet->getHighestColumn() . $myRow
//    );


exit;
//setCellColor($objPHPExcel,'A1','FF0000');

//$styleArray = array(
//    'font'  => array(
//        'color' => array('rgb' => '000000'),
//        'name'  => 'Verdana'
//    ));
//
//$worksheet->getStyle('A1')->applyFromArray($styleArray);

//$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
//$objWriter->save($file);
exit;