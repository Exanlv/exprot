<?php
/**
 * @var \Exan\Exprot\SheetInterface[] $sheets
 */
?>
<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="fntdata" ContentType="application/x-fontdata" />
    <Default Extension="jpeg" ContentType="image/jpeg" />
    <Default Extension="png" ContentType="image/png" />
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
    <Default Extension="xml" ContentType="application/xml" />
    <Override PartName="/xl/workbook.xml"
        ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />

    <?php foreach ($sheets as $sheet): ?>
        <Override PartName="/xl/worksheets/<?= $sheet->getSlug() ?>.xml"
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
    <?php endforeach ?>
    <Override PartName="/docProps/core.xml"
        ContentType="application/vnd.openxmlformats-package.core-properties+xml" />
    <Override PartName="/docProps/app.xml"
        ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml" />
</Types>
