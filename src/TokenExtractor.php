<?php

namespace Application;

class TokenExtractor
{


    private $nicknamesColumn = 'D';
    private $tokensColumn = 'E';
    /** @var \PHPExcel_Worksheet */
    private $worksheet;

    /**
     * Updater constructor.
     */
    public function __construct($file, $lmEditionNumber)
    {

        /** @var \PHPExcel_Reader_Excel2007 $reader */
        $reader = new \PHPExcel_Reader_Excel2007();
        $reader->setLoadSheetsOnly($lmEditionNumber);
        $objPHPExcel = $reader->load($file);
        if (!$objPHPExcel->getSheetByName($lmEditionNumber)) {
            throw new \Exception(sprintf("No sheet %s in file %s", $lmEditionNumber, $file));
        }
        $this->worksheet = $objPHPExcel->getActiveSheet();


    }


    public function getResultsArray()
    {
        $highestRow = $this->worksheet->getHighestRow();

        $results = [];

        for ($row = 2; $row <= $highestRow; $row++) {

            $tokens = $this->worksheet->getCell($this->tokensColumn . $row)->getFormattedValue();
            if (!$tokens) {
                continue;
            }
            $nickname = $this->worksheet->getCell($this->nicknamesColumn . $row)->getFormattedValue();
            $results[$nickname] = $tokens;
        }
        return $results;
    }


}