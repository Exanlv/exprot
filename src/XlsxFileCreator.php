<?php

declare(strict_types=1);

namespace Exan\Exprot;

use League\Plates\Engine;
use RuntimeException;

class XlsxFileCreator
{
    private array $sheets = [];

    public function __construct(
        public readonly string $tmpDir,
        public readonly Engine $templatingEngine,
    ) {
    }

    public function addSheet(string $name, Sheet $sheet)
    {
        $this->sheets[$name] = $sheet;
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
            'xl/workbook',
            [
                'sheets' => array_keys($this->sheets),
            ],
        );

        $this->createFile(
            $dir,
            '[Content_Types]',
            [
                'sheets' => array_keys($this->sheets),
            ],
        );

        $this->createFile(
            $dir,
            'docProps/app',
            [
                'sheets' => array_keys($this->sheets),
            ],
        );

        $this->createFile(
            $dir,
            'docProps/core',
        );

        $this->createFile(
            $dir,
            'xl/theme/theme1',
        );

        $this->createFile(
            $dir,
            'xl/styles',
        );

        $this->createFile(
            $dir,
            'xl/worksheets/sheet1',
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
                'sheets' => array_keys($this->sheets),
            ],
            'workbook.xml.rels',
        );
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
