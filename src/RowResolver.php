<?php

declare(strict_types=1);

namespace Exan\Exprot;


interface RowResolver
{
    /** @return Array<Array<string>> */
    public function getRowBatch(int $i): ?array;
}
