<?php

declare(strict_types=1);

namespace Exan\Exprot;

class Excel
{
    private array $sheets = [];

    public function __construct(
        public readonly string $outFile,
    ) {
    }

    public function addSheet(string $name, SheetWriter $sheet)
    {
        $this->sheets[$name] = $sheet;
    }

    public function create()
    {

    }
}
