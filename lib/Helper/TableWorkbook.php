<?php

declare(strict_types=1);

namespace Slam\Excel\Helper;

use Slam\Excel;

final class TableWorkbook extends Excel\Pear\Writer\Workbook
{
    public const GREY_MEDIUM   = 43;
    public const GREY_LIGHT    = 42;

    /**
     * @var int
     */
    private $rowsPerSheet = 60000;

    /**
     * @var string
     */
    private $emptyTableMessage = '';

    /**
     * @var CellStyle\Text
     */
    private $styleIdentity;

    /**
     * @var null|array
     */
    private $formats;

    public function __construct(string $filename)
    {
        parent::__construct($filename);

        $this->setCustomColor(self::GREY_MEDIUM,    0xCC, 0xCC, 0xCC);
        $this->setCustomColor(self::GREY_LIGHT,     0xE8, 0xE8, 0xE8);

        $this->styleIdentity = new CellStyle\Text();
    }

    public function setRowsPerSheet(int $rowsPerSheet): void
    {
        $this->rowsPerSheet = $rowsPerSheet;
    }

    public function setEmptyTableMessage(string $emptyTableMessage): void
    {
        $this->emptyTableMessage = $emptyTableMessage;
    }

    public static function getColumnStringFromIndex(int $index): string
    {
        if ($index < 0) {
            throw new Excel\Exception\InvalidArgumentException('Column index must be equal or greater than zero');
        }

        static $indexCache = [];

        if (! isset($indexCache[$index])) {
            if ($index < 26) {
                $indexCache[$index] = \chr(65 + $index);
            } elseif ($index < 702) {
                $indexCache[$index] = \chr(64 + (int) ($index / 26))
                    . \chr(65 + $index % 26)
                ;
            } else {
                $indexCache[$index] = \chr(64 + (int) (($index - 26) / 676))
                    . \chr(65 + (int) ((($index - 26) % 676) / 26))
                    . \chr(65 + $index % 26)
                ;
            }
        }

        return $indexCache[$index];
    }

    public function writeTable(Table $table): Table
    {
        $this->writeTableHeading($table);
        $tables = [$table];

        $count      = 0;
        $headingRow = true;
        foreach ($table->getData() as $row) {
            ++$count;

            if ($table->getRowCurrent() >= $this->rowsPerSheet) {
                $table    = $table->splitTableOnNewWorksheet($this->addWorksheet(\uniqid()));
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

        if (\count($tables) > 1) {
            \reset($tables);
            $table      = \current($tables);
            $firstSheet = $table->getActiveSheet();
            // In Excel the maximum length for a sheet name is 30
            $originalName = \mb_substr($firstSheet->name, 0, 21);

            $sheetCounter = 0;
            $sheetTotal   = \count($tables);
            foreach ($tables as $table) {
                ++$sheetCounter;
                $table->getActiveSheet()->name = \sprintf('%s (%s|%s)', $originalName, $sheetCounter, $sheetTotal);
            }
        }

        if ($table->getFreezePanes()) {
            foreach ($tables as $table) {
                $table->getActiveSheet()->freezePanes([$table->getRowStart() + 2, 0]);
            }
        }

        if (0 === $count) {
            $table->incrementRow();
            $table->getActiveSheet()->writeString($table->getRowCurrent(), $table->getColumnCurrent(), $this->emptyTableMessage);
            $table->incrementRow();
        }

        $table->setCount($count);

        \end($tables);

        return \current($tables);
    }

    private function writeTableHeading(Table $table): void
    {
        $table->resetColumn();
        $table->getActiveSheet()->writeString($table->getRowCurrent(), $table->getColumnCurrent(), $this->sanitize($table->getHeading()));
        $table->incrementRow();
    }

    private function writeColumnsHeading(Table $table, array $row): void
    {
        $columnCollection = $table->getColumnCollection();
        $columnKeys       = \array_keys($row);
        $this->generateFormats($table, $columnKeys, $columnCollection);

        $table->resetColumn();
        $titles = [];
        foreach ($columnKeys as $title) {
            $width    = 10;
            $newTitle = \ucwords(\str_replace('_', ' ', $title));

            if (isset($columnCollection) && isset($columnCollection[$title])) {
                $width    = $columnCollection[$title]->getWidth();
                $newTitle = $columnCollection[$title]->getHeading();
            }

            $table->getActiveSheet()->setColumn($table->getColumnCurrent(), $table->getColumnCurrent(), $width);
            $titles[$title] = $newTitle;

            $table->incrementColumn();
        }

        $this->writeRow($table, $titles, 'title');

        $table->setWrittenColumnTitles($titles);
        $table->flagDataRowStart();
    }

    private function writeRow(Table $table, array $row, ?string $type = null): void
    {
        $table->resetColumn();
        $sheet = $table->getActiveSheet();

        foreach ($row as $key => $content) {
            $cellStyle = $this->styleIdentity;
            $format    = null;
            if (isset($this->formats[$key])) {
                if (null === $type) {
                    $type = (($table->getRowCurrent() % 2)
                        ? 'zebra_light'
                        : 'zebra_dark'
                    );
                }
                $cellStyle = $this->formats[$key]['cell_style'];
                $format    = $this->formats[$key][$type];
            }

            $write = 'write';
            if (\get_class($cellStyle) === \get_class($this->styleIdentity)) {
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

    /**
     * @param mixed $value
     */
    private function sanitize($value): string
    {
        static $sanitizeMap = [
            '&amp;'     => '&',
            '&lt;'      => '<',
            '&gt;'      => '>',
            '&apos;'    => '\'',
            '&quot;'    => '"',
        ];

        $value = \str_replace(
            \array_keys($sanitizeMap),
            \array_values($sanitizeMap),
            (string) $value
        );
        $value = \mb_convert_encoding($value, 'Windows-1252');

        return $value;
    }

    private function generateFormats(Table $table, array $titles, ?ColumnCollectionInterface $columnCollection = null): void
    {
        $this->formats = [];
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

            $zebraDark = $this->addFormat();
            $zebraDark->setColor('black');
            $zebraDark->setSize($table->getFontSize());
            $zebraDark->setFgColor(self::GREY_LIGHT);

            if ($table->getTextWrap()) {
                $zebraLight->setTextWrap();
                $zebraLight->setAlign('top');

                $zebraDark->setTextWrap();
                $zebraDark->setAlign('top');
            }

            $this->formats[$key] = [
                'cell_style'    => null,
                'title'         => $header,
                'zebra_dark'    => $zebraLight,
                'zebra_light'   => $zebraDark,
            ];

            $cellStyle = $this->styleIdentity;
            if (isset($columnCollection) && isset($columnCollection[$key])) {
                $cellStyle = $columnCollection[$key]->getCellStyle();
            }

            $cellStyle->styleCell($zebraLight);
            $cellStyle->styleCell($zebraDark);

            $this->formats[$key]['cell_style'] = $cellStyle;
        }
    }
}
