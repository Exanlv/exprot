<?php

use Exan\Exprot\RowResolver;
use Exan\Exprot\Sheet;
use Exan\Exprot\SheetWriter;
use Exan\Exprot\XlsxFileCreator;
use League\Plates\Engine;

require './vendor/autoload.php';

$resolver = new class implements RowResolver {
    private readonly PDO $connection;

    public function __construct()
    {
        $this->connection = new PDO('sqlite:./source.db');
    }

    public function getRowBatch(int $i): ?array
    {
        dump($i);
        $limit = 1000;
        $offset = $limit * $i;

        $stmt = $this->connection->prepare('SELECT * FROM users LIMIT :limit OFFSET :offset');
        $stmt->bindValue(':limit', $limit, PDO::PARAM_INT);
        $stmt->bindValue(':offset', $offset, PDO::PARAM_INT);
        $stmt->execute();

        $users = $stmt->fetchAll(PDO::FETCH_ASSOC);

        if (count($users) === 0) {
            return null;
        }

        $data = [];

        foreach ($users as $record) {
            $data[] = array_map(
                fn (mixed $data) => (string) $data,
                array_values($record),
            );
        }

        return $data;
    }
};

$excel = new XlsxFileCreator('./tmp', new Engine(__DIR__ . '/resources/excel'));

$excel->addSheet(new Sheet('rId3', 'Sheet 1', 'sheet1', new SheetWriter($resolver)));
// $excel->addSheet(new Sheet('rId4', 'Sheet 2', 'sheet2', new SheetWriter($resolver)));

$excel->create('test');
