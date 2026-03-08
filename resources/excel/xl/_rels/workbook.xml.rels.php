<?php
/**
 * @var \Exan\Exprot\Sheet[] $sheets
 */
?>
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <?php foreach ($sheets as $sheet): ?>
    <Relationship Id="<?= $sheet->rId ?>" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/<?= $sheet->slug ?>.xml" />
    <?php endforeach; ?>
</Relationships>
