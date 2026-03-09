<?php

declare(strict_types=1);

namespace Exan\Exprot;

use League\Plates\Engine;
use RuntimeException;
use ZipArchive;

class XlsxFileCreator
{
    /** @var Sheet[] */
    private array $sheets = [];

    public function __construct(
        public readonly string $tmpDir,
        public readonly Engine $templatingEngine,
    ) {}

    public function addSheet(Sheet $sheet)
    {
        $this->sheets[] = $sheet;
    }

    public function create(string $outFile)
    {
        $slug = str_replace(['/', '\\', DIRECTORY_SEPARATOR], '_', $outFile);
        $dir = $this->tmpDir . '/' . $slug;

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

        $this->createZip($outFile, $dir);

        $this->rmDir($dir);
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

    private function createZip(string $outFile, string $dir)
    {
        $zip = new ZipArchive();

        $opened = $zip->open($outFile, ZipArchive::CREATE | ZipArchive::OVERWRITE);
        if ($opened !== true) {
            throw new RuntimeException(sprintf('Unable to create ZipArchive at %s, error code %d', $outFile, $opened));
        }

        $this->addFilesToArchive($zip, $dir);

        $zip->close();
    }

    private function addFilesToArchive(ZipArchive $zip, string $dir, string $basePath = ''): void
    {
        $files = scandir($dir);

        foreach ($files as $file) {
            if ($file === '.' || $file === '..') {
                continue;
            }

            $fullPath = $dir . '/' . $file;
            $relativePath = $basePath . $file;

            if (is_file($fullPath)) {
                $zip->addFile($fullPath, $relativePath);
            } elseif (is_dir($fullPath)) {
                $this->addFilesToArchive($zip, $fullPath, $relativePath . '/');
            }
        }
    }

    private function rmDir(string $dir)
    {
        if (!is_dir($dir)) {
            return false;
        }

        $items = scandir($dir);
        foreach ($items as $item) {
            if ($item === '.' || $item === '..') {
                continue;
            }

            $path = $dir . DIRECTORY_SEPARATOR . $item;
            if (is_dir($path)) {
                $this->rmDir($path);
            } else {
                unlink($path);
            }
        }

        return rmdir($dir);
    }
}
