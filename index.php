<?php

use Exan\Exprot\Excel;
use Exan\Exprot\RowResolver;

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

$excel = new Excel('out.xlsx', $resolver);
$excel->setTmpDir('./tmp');
$excel->write();
