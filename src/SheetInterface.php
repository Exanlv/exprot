<?php

namespace Exan\Exprot;

use Generator;

interface SheetInterface
{
    public function getName(): string;
    public function getSlug(): string;
    /** @return string[] */
    public function getHeaders(): array;
    public function getRows(): Generator;
}
