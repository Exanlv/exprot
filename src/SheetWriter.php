<?php

declare(strict_types=1);

namespace Exan\Exprot;

class SheetWriter
{
    public function __construct(
        private readonly SheetInterface $sheet,
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
        $this->writeHeaders($file);
        $this->writeRows($file);
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
    private function writeHeaders(mixed $file)
    {
        fwrite($file, $this->getRow($this->sheet->getHeaders(), 0, true));
    }

    /** @param resource $file */
    private function writeRows(mixed $file)
    {
        foreach ($this->sheet->getRows() as $i => $row) {
            fwrite($file, $this->getRow($row, $i + 1));
        }
    }

    private function getRow(array $row, int $rowNumber, bool $isHeader = false): string
    {
        $rowi1 = $rowNumber + 1;

        $rpr = $isHeader
            ? '<rPr><i/></rPr>'
            : '';

        $rowXml = implode(
            '',
            array_map(
                function (?string $value, int $colNumber) use ($rowNumber, $rpr) {
                    $value = htmlspecialchars($value ?? '', ENT_XML1 | ENT_QUOTES, 'UTF-8');
                    $fullColName = Helper::indexToExcelColumn($colNumber) . ($rowNumber + 1);

                    if ($rpr) {
                        return "<c r=\"$fullColName\" t=\"inlineStr\"><is><r>$rpr<t>$value</t></r></is></c>";
                    }

                    return "<c r=\"$fullColName\" t=\"inlineStr\"><is><t>$value</t></is></c>";
                },
                $row,
                range(0, count($row) - 1),
            )
        );

        return "<row r=\"$rowi1\">" . $rowXml . "</row>";
    }
}
