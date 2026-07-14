<?php
/**
 * @var \Exan\Exprot\SheetInterface[] $sheets
 */
?>
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <?php foreach ($sheets as $i => $sheet): ?>
    <Relationship Id="rId<?= $i + 1 ?>" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/<?= $sheet->getSlug() ?>.xml" />
    <?php endforeach; ?>
</Relationships>
