<?php

use Exan\Exprot\SheetInterface;
use Exan\Exprot\XlsxFileCreator;
use League\Plates\Engine;

require './vendor/autoload.php';

class MySheet implements SheetInterface
{
    public function __construct(
        private readonly string $name,
        private readonly string $slug,
        private readonly array $headers,
    ) {
    }

    public function getName(): string
    {
        return $this->name;
    }

    public function getSlug(): string
    {
        return $this->slug;
    }

    #[Override]
    public function getHeaders(): array
    {
        return $this->headers;
    }

    public function getRows(): Generator
    {
        for ($index = 1; $index <= 50; $index++) {
            yield array_map(
                fn (string $header) => sprintf('%s %d', $header, $index),
                $this->headers
            );
        }
    }
}

$excel = new XlsxFileCreator(__DIR__ . '/kaas', new Engine(__DIR__ . '/resources/excel'));

$excel->addSheet(new MySheet('Marketing Leads', 'marketing-leads', ['Name', 'Company', 'Söurce', 'Status', 'Assigned To']));
$excel->addSheet(new MySheet('Finance Report', 'finance-report', ['Account', 'Category', 'Amoïnt', 'Currency', 'Date']));
$excel->addSheet(new MySheet('Demo Data', 'demo-data', ['Foo', 'Bar', 'Baz']));

$excel->create('out.xlsx');

dump(memory_get_peak_usage(true));
