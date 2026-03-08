<?php

use Exan\Exprot\RowResolver;
use Exan\Exprot\Sheet;
use Exan\Exprot\SheetWriter;
use Exan\Exprot\XlsxFileCreator;
use League\Plates\Engine;

require './vendor/autoload.php';

$resolver = new class implements RowResolver {
    public function getRowBatch(int $i): ?array
    {
        if ($i > 10) {
            return null;
        }

        $batch = [];
        for ($rowNum = 1; $rowNum <= 5; $rowNum++) {
            $batch[] = [
                "row " . (($i * 5) + $rowNum),                  // row number
                "row " . (($i * 5) + $rowNum) . " value 2",
                "row " . (($i * 5) + $rowNum) . " value 3",
                "row " . (($i * 5) + $rowNum) . " value 4",
            ];
        }

        return $batch;
    }
};

$excel = new XlsxFileCreator('./tmp', new Engine(__DIR__ . '/resources/excel'));

$excel->addSheet(new Sheet('rId3', 'Sheet 1', 'sheet1', new SheetWriter($resolver)));

$excel->create('test');
