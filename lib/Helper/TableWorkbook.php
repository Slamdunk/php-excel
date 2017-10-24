<?php

declare(strict_types=1);

namespace Slam\Excel\Helper;

use Slam\Excel;

final class TableWorkbook extends Excel\Pear\Writer\Workbook
{
    const GREY_DARK     = 60;
    const GREY_MEDIUM   = 61;
    const GREY_LIGHT    = 62;

    private $rowsPerSheet = 60000;

    private $emptyTableMessage = '';

    private $styleIdentity;

    private $formats;

    public function __construct(string $filename)
    {
        parent::__construct($filename);

        $this->setCustomColor(self::GREY_DARK,      hexdec('7f'), hexdec('7f'), hexdec('7f'));
        $this->setCustomColor(self::GREY_MEDIUM,    hexdec('cc'), hexdec('cc'), hexdec('cc'));
        $this->setCustomColor(self::GREY_LIGHT,     hexdec('e8'), hexdec('e8'), hexdec('e8'));

        $this->styleIdentity = new Excel\Helper\CellStyle\Text();
    }

    public function setRowsPerSheet(int $rowsPerSheet)
    {
        $this->rowsPerSheet = $rowsPerSheet;
    }

    public function setEmptyTableMessage(string $emptyTableMessage)
    {
        $this->emptyTableMessage = $emptyTableMessage;
    }

    public function writeTable(Table $table): Table
    {
        $this->writeTableHeading($table);
        $tables = array($table);

        $count = 0;
        $headingRow = true;
        foreach ($table->getData() as $row) {
            ++$count;

            if ($table->getRowCurrent() >= $this->rowsPerSheet) {
                $table = $table->splitTableOnNewWorksheet($this->addWorksheet(uniqid()));
                $tables[] = $table;
                $this->writeTableHeading($table);
                $headingRow = true;
            }

            if ($headingRow) {
                $this->writeColumnsHeading($table, $row);

                $headingRow = false;
            }

            $this->writeRow($table, $row);
        }

        if (count($tables) > 1) {
            $table = reset($tables);
            $firstSheet = $table->getActiveSheet();
            // In Excel the maximum length for a sheet name is 30
            $originalName = mb_substr($firstSheet->name, 0, 21);

            $sheetCounter = 0;
            $sheetTotal = count($tables);
            foreach ($tables as $table) {
                ++$sheetCounter;
                $table->getActiveSheet()->name = sprintf('%s (%s|%s)', $originalName, $sheetCounter, $sheetTotal);
            }
        }

        if ($table->getFreezePanes()) {
            foreach ($tables as $table) {
                $table->getActiveSheet()->freezePanes(array($table->getRowStart() + 2, 0));
            }
        }

        if (0 === $count) {
            $table->incrementRow();
            $table->getActiveSheet()->writeString($table->getRowCurrent(), $table->getColumnCurrent(), $this->emptyTableMessage);
            $table->incrementRow();
        }

        $table->setCount($count);

        return end($tables);
    }

    private function writeTableHeading(Table $table)
    {
        $table->resetColumn();
        $table->getActiveSheet()->writeString($table->getRowCurrent(), $table->getColumnCurrent(), $this->sanitize($table->getHeading()));
        $table->incrementRow();
    }

    private function writeColumnsHeading(Table $table, array $row)
    {
        $columnCollection = $table->getColumnCollection();
        $columnKeys = array_keys($row);
        $this->generateFormats($table, $columnKeys, $columnCollection);

        $table->resetColumn();
        $titles = array();
        foreach ($columnKeys as $title) {
            $width = 10;
            $newTitle = ucwords(str_replace('_', ' ', $title));

            if (isset($columnCollection) and isset($columnCollection[$title])) {
                $width = $columnCollection[$title]->getWidth();
                $newTitle = $columnCollection[$title]->getHeading();
            }

            $table->getActiveSheet()->setColumn($table->getColumnCurrent(), $table->getColumnCurrent(), $width);
            $titles[$title] = $newTitle;

            $table->incrementColumn();
        }

        $this->writeRow($table, $titles, 'title');
    }

    private function writeRow(Table $table, array $row, string $type = null)
    {
        $table->resetColumn();
        $sheet = $table->getActiveSheet();

        foreach ($row as $key => $content) {
            $cellStyle = $this->styleIdentity;
            $format = null;
            if (isset($this->formats[$key])) {
                if (null === $type) {
                    $type = (($table->getRowCurrent() % 2)
                        ? 'zebra_dark'
                        : 'zebra_light'
                    );
                }
                $cellStyle = $this->formats[$key]['cell_style'];
                $format = $this->formats[$key][$type];
            }

            $write = 'write';
            if (get_class($cellStyle) === get_class($this->styleIdentity)) {
                $write = 'writeString';
            }

            $content = $cellStyle->decorateValue($content);
            $content = $this->sanitize($content);

            $sheet->{$write}($table->getRowCurrent(), $table->getColumnCurrent(), $content, $format);

            $table->incrementColumn();
        }

        if (null !== ($rowHeight = $table->getRowHeight())) {
            $sheet->setRow($table->getRowCurrent(), $rowHeight);
        }

        $table->incrementRow();
    }

    private function sanitize($value)
    {
        static $sanitizeMap;

        if (null === $sanitizeMap) {
            $sanitizeMap = array(
                '&amp;'     => '&',
                '&lt;'      => '<',
                '&gt;'      => '>',
                '&apos;'    => "'",
                '&quot;'    => '"',
            );
        }

        $value = str_replace(
            array_keys($sanitizeMap),
            array_values($sanitizeMap),
            $value
        );
        $value = mb_convert_encoding($value, 'Windows-1252');

        return $value;
    }

    private function generateFormats(Table $table, array $titles, ColumnCollectionInterface $columnCollection = null)
    {
        $this->formats = array();
        foreach ($titles as $key) {
            $header = $this->addFormat();
            $header->setColor('black');
            $header->setSize($table->getFontSize());
            $header->setBold();
            $header->setFgColor(self::GREY_MEDIUM);
            $header->setTextWrap();
            $header->setAlign('center');

            $zebraLight = $this->addFormat();
            $zebraLight->setColor('black');
            $zebraLight->setSize($table->getFontSize());
            $zebraLight->setFgColor('white');
            $zebraLight->SetBorderColor(self::GREY_DARK);

            $zebraDark = $this->addFormat();
            $zebraDark->setColor('black');
            $zebraDark->setSize($table->getFontSize());
            $zebraDark->setFgColor(self::GREY_LIGHT);
            $zebraDark->SetBorderColor(self::GREY_DARK);

            if ($table->getTextWrap()) {
                $zebraLight->setTextWrap();
                $zebraLight->setAlign('top');

                $zebraDark->setTextWrap();
                $zebraDark->setAlign('top');
            }

            $this->formats[$key] = array(
                'cell_style'    => null,
                'title'         => $header,
                'zebra_dark'    => $zebraLight,
                'zebra_light'   => $zebraDark,
            );

            $cellStyle = $this->styleIdentity;
            if (isset($columnCollection) and isset($columnCollection[$key])) {
                $cellStyle = $columnCollection[$key]->getCellStyle();
            }

            $cellStyle->styleCell($zebraLight);
            $cellStyle->styleCell($zebraDark);

            $this->formats[$key]['cell_style'] = $cellStyle;
        }
    }
}
