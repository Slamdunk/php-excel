<?php

declare(strict_types=1);

namespace Slam\Excel\Helper;

use Countable;
use Iterator;
use Slam\Excel\Exception;
use Slam\Excel\Pear\Writer\Worksheet;

final class Table implements Countable
{
    /**
     * @var Worksheet
     */
    private $activeSheet;

    /**
     * @var null|int
     */
    private $dataRowStart;

    /**
     * @var int
     */
    private $rowStart;

    /**
     * @var int
     */
    private $rowEnd;

    /**
     * @var int
     */
    private $rowCurrent;

    /**
     * @var int
     */
    private $columnStart;

    /**
     * @var int
     */
    private $columnEnd;

    /**
     * @var int
     */
    private $columnCurrent;

    /**
     * @var string
     */
    private $heading;

    /**
     * @var Iterator
     */
    private $data;

    /**
     * @var null|ColumnCollectionInterface
     */
    private $columnCollection;

    /**
     * @var bool
     */
    private $freezePanes = true;

    /**
     * @var int
     */
    private $fontSize = 8;

    /**
     * @var null|int
     */
    private $rowHeight;

    /**
     * @var bool
     */
    private $textWrap = false;

    /**
     * @var null|array
     */
    private $writtenColumnTitles;

    /**
     * @var null|int
     */
    private $count;

    public function __construct(Worksheet $activeSheet, int $row, int $column, string $heading, Iterator $data)
    {
        $this->activeSheet = $activeSheet;

        $this->rowStart   =
        $this->rowEnd     =
        $this->rowCurrent =
            $row
        ;

        $this->columnStart   =
        $this->columnEnd     =
        $this->columnCurrent =
            $column
        ;

        $this->heading = $heading;

        $this->data = $data;
    }

    public function getActiveSheet(): Worksheet
    {
        return $this->activeSheet;
    }

    public function getDataRowStart(): ?int
    {
        return $this->dataRowStart;
    }

    public function flagDataRowStart(): void
    {
        $this->dataRowStart = $this->rowCurrent;
    }

    public function getRowStart(): int
    {
        return $this->rowStart;
    }

    public function getRowEnd(): int
    {
        return $this->rowEnd;
    }

    public function getRowCurrent(): int
    {
        return $this->rowCurrent;
    }

    public function incrementRow(): void
    {
        ++$this->rowCurrent;

        $this->rowEnd = \max($this->rowEnd, $this->rowCurrent);
    }

    public function getColumnStart(): int
    {
        return $this->columnStart;
    }

    public function getColumnEnd(): int
    {
        return $this->columnEnd;
    }

    public function getColumnCurrent(): int
    {
        return $this->columnCurrent;
    }

    public function incrementColumn(): void
    {
        ++$this->columnCurrent;

        $this->columnEnd = \max($this->columnEnd, $this->columnCurrent);
    }

    public function resetColumn(): void
    {
        $this->columnCurrent = $this->columnStart;
    }

    public function getHeading(): string
    {
        return $this->heading;
    }

    public function getData(): Iterator
    {
        return $this->data;
    }

    public function setColumnCollection(?ColumnCollectionInterface $columnCollection): void
    {
        $this->columnCollection = $columnCollection;
    }

    public function getColumnCollection(): ?ColumnCollectionInterface
    {
        return $this->columnCollection;
    }

    public function setFreezePanes(bool $freezePanes): void
    {
        $this->freezePanes = $freezePanes;
    }

    public function getFreezePanes(): bool
    {
        return $this->freezePanes;
    }

    public function setFontSize(int $fontSize): void
    {
        $this->fontSize = $fontSize;
    }

    public function getFontSize(): int
    {
        return $this->fontSize;
    }

    public function setRowHeight(?int $rowHeight): void
    {
        $this->rowHeight = $rowHeight;
    }

    public function getRowHeight(): ?int
    {
        return $this->rowHeight;
    }

    public function setTextWrap(bool $textWrap): void
    {
        $this->textWrap = $textWrap;
    }

    public function getTextWrap(): bool
    {
        return $this->textWrap;
    }

    public function setWrittenColumnTitles(?array $writtenColumnTitles): void
    {
        $this->writtenColumnTitles = $writtenColumnTitles;
    }

    public function getWrittenColumnTitles(): ?array
    {
        return $this->writtenColumnTitles;
    }

    public function setCount(int $count): void
    {
        $this->count = $count;
    }

    /**
     * @return null|int
     */
    public function count()
    {
        if (null === $this->count) {
            throw new Exception\RuntimeException('Workbook must set count on table');
        }

        return $this->count;
    }

    public function isEmpty(): bool
    {
        return 0 === $this->count();
    }

    public function splitTableOnNewWorksheet(Worksheet $activeSheet): self
    {
        $newTable = new self($activeSheet, 0, $this->getColumnStart(), $this->getHeading(), $this->getData());
        $newTable->setColumnCollection($this->getColumnCollection());
        $newTable->setFreezePanes($this->getFreezePanes());

        return $newTable;
    }
}
