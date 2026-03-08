<?php

declare(strict_types=1);

namespace Exan\Exprot;

use League\Plates\Engine;
use RuntimeException;

class XlsxFileCreator
{
    /** @var Sheet[] */
    private array $sheets = [];

    public function __construct(
        public readonly string $tmpDir,
        public readonly Engine $templatingEngine,
    ) {
    }

    public function addSheet(Sheet $sheet)
    {
        $this->sheets[] = $sheet;
    }

    public function create(string $outFile)
    {
        $dir = $this->tmpDir . DIRECTORY_SEPARATOR . $outFile;

        if (is_dir($dir)) {
            throw new RuntimeException(sprintf('Unable to create working dir "%s", directory already exists', $dir));
        }

        mkdir($dir, recursive: true);

        $this->createFile(
            $dir,
            '[Content_Types]',
            [
                'sheets' => $this->sheets,
            ],
        );

        $this->createFile(
            $dir,
            '_rels/.rels',
            as: '.rels',
        );

        $this->createFile(
            $dir,
            'xl/_rels/workbook.xml.rels',
            [
                'sheets' => $this->sheets,
            ],
            'workbook.xml.rels',
        );

        $this->createFile(
            $dir,
            'xl/workbook',
            [
                'sheets' => $this->sheets,
            ],
        );

        $this->createFile(
            $dir,
            'docProps/app',
            [
                'sheets' => $this->sheets,
            ],
        );

        $this->createFile(
            $dir,
            'docProps/core',
        );

        foreach ($this->sheets as $sheet) {
            $sheet->writer->write($dir . '/xl/worksheets', $sheet->slug . '.xml');
        }
    }

    private function createFile(string $workDir, string $template, array $variables = [], ?string $as = null): void
    {
        $path = explode('/', $template);
        $fileName = array_pop($path) . '.xml';

        if ($as) {
            $fileName = $as;
        }

        $dirPath = $workDir . DIRECTORY_SEPARATOR . implode(DIRECTORY_SEPARATOR, $path);

        if (!is_dir($dirPath)) {
            mkdir($dirPath, recursive: true);
        }

        file_put_contents(
            $dirPath . DIRECTORY_SEPARATOR . $fileName,
            $this->templatingEngine->render($template, $variables)
        );
    }
}
