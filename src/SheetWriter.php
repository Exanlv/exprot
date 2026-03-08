<?php

declare(strict_types=1);

namespace Exan\Exprot;

class SheetWriter
{
    public function __construct(
        private readonly RowResolver $rowResolver,
    ) {
    }

    public function write(string $directory, string $fileName): void
    {
        if (!is_dir($directory)) {
            mkdir($directory, recursive: true);
        }

        $outFile = $directory . DIRECTORY_SEPARATOR . $fileName;


        $file = fopen($outFile, 'w');

        $this->startFile($file);
        $this->writeRowBatchFiles($file);
        $this->endFile($file);
    }

    /** @param resource $file */
    private function startFile(mixed $file): void
    {
        fwrite($file, <<<XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
            xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
            xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
            <sheetData>

        XML);
    }

    /** @param resource $file */
    public function endFile(mixed $file): void
    {
        fwrite($file, <<<XML

            </sheetData>
        </worksheet>
        XML);
    }

    /** @param resource $file */
    private function writeRowBatchFiles(mixed $file)
    {
        $rowCount = 0;
        $i = 0;
        $rows = $this->rowResolver->getRowBatch($i);

        while ($rows !== null) {
            $this->writeRowBatch($file, $rows, $i, $rowCount);
            $rowCount += count($rows);

            $i++;
            $rows = $this->rowResolver->getRowBatch($i);
        }
    }

    /** @param resource $file */
    private function writeRowBatch(mixed $file, array $batch, int $batchNumber, int $baseRowNumber): void
    {
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
}
