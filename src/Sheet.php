<?php

declare(strict_types=1);

namespace Exan\Exprot;

readonly class Sheet
{
    public function __construct(
        public string $rId,
        public string $name,
        public string $slug,
        public SheetWriter $writer,
    ) {
    }
}
