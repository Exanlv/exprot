<?php

declare(strict_types=1);

namespace Exan\Exprot;

class SheetWriter
{
    private string $tmpDir;

    public function __construct(
        private readonly RowResolver $rowResolver,
    ) {
        $this->tmpDir = trim(sys_get_temp_dir(), '/\\');
    }

    public function setTmpDir(string $path)
    {
        $this->tmpDir = $path;
    }

    public function write(): void
    {
        $this->prepFile();
        $this->writeRowBatchFiles();
        $this->combineBatchFiles();
        $this->zip();
    }

    private function prepFile()
    {

    }

    private function writeRowBatchFiles()
    {
        $rowCount = 0;
        $i = 0;
        $rows = $this->rowResolver->getRowBatch($i);

        while ($rows !== null) {
            $this->writeRowBatch($rows, $i, $rowCount);
            $rowCount += count($rows);

            $i++;
            $rows = $this->rowResolver->getRowBatch($i);
        }
    }

    private function writeRowBatch(array $batch, int $batchNumber, int $baseRowNumber): void
    {
        $fileName = $this->tmpDir . '/' . $this->outFile . '-batch-' . ((int) $batchNumber);
        $file = fopen($fileName, 'w');

        $batch = array_values($batch);

        foreach ($batch as $i => $row) {
            fwrite($file, $this->getRow($row, $baseRowNumber + $i));
        }
    }

    private function getRow(array $row, int $rowNumber): string
    {
        $rowi1 = $rowNumber + 1;

        $rowXml = implode(
            '',
            array_map(
                function (string $value, int $colNumber) use ($rowNumber) {
                    $fullColName = Helper::indexToExcelColumn($colNumber) . ($rowNumber + 1);
                    return "<c r=\"$fullColName\" t=\"inlineStr\"><is><t>$value</t></is></c>";
                },
                $row,
                range(0, count($row) - 1),
            )
        );

        return "<row r=\"$rowi1\">" . $rowXml . "</row>";
    }

    private function combineBatchFiles()
    {

    }

    private function zip()
    {

    }
}
