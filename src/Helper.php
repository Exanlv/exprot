<?php

declare(strict_types=1);

namespace Exan\Exprot;

class Helper
{
    public static function indexToExcelColumn(int $index): string
    {
        $index++;
        $column = '';

        while ($index > 0) {
            $remainder = ($index - 1) % 26;
            $column = chr(65 + $remainder) . $column;
            $index = intdiv($index - 1, 26);
        }

        return $column;
    }
}
