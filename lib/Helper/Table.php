<?php

declare(strict_types=1);

namespace Slam\Excel\Helper;

use Countable;
use Iterator;
use Slam\Excel\Exception;
use Slam\Excel\Pear\Writer\Worksheet;

final class Table implements Countable
{
    private $activeSheet;

    private $rowStart;
    private $rowEnd;
    private $rowCurrent;

    private $columnStart;
    private $columnEnd;
    private $columnCurrent;

    private $heading;

    private $data;

    private $columnCollection;

    private $freezePanes = true;
    private $fontSize = 8;
    private $rowHeight;
    private $textWrap = false;

    private $count;

    public function __construct(Worksheet $activeSheet, int $row, int $column, string $heading, Iterator $data)
    {
        $this->activeSheet = $activeSheet;

        $this->rowStart =
        $this->rowEnd =
        $this->rowCurrent =
            $row
        ;

        $this->columnStart =
        $this->columnEnd =
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

    public function incrementRow()
    {
        ++$this->rowCurrent;

        $this->rowEnd = max($this->rowEnd, $this->rowCurrent);
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

    public function incrementColumn()
    {
        ++$this->columnCurrent;

        $this->columnEnd = max($this->columnEnd, $this->columnCurrent);
    }

    public function resetColumn()
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

    public function setColumnCollection(ColumnCollectionInterface $columnCollection = null)
    {
        $this->columnCollection = $columnCollection;

        return $this;
    }

    public function getColumnCollection()
    {
        return $this->columnCollection;
    }

    public function setFreezePanes(bool $freezePanes)
    {
        $this->freezePanes = $freezePanes;

        return $this;
    }

    public function getFreezePanes(): bool
    {
        return $this->freezePanes;
    }

    public function setFontSize(int $fontSize)
    {
        $this->fontSize = $fontSize;

        return $this;
    }

    public function getFontSize(): int
    {
        return $this->fontSize;
    }

    public function setRowHeight(int $rowHeight = null)
    {
        $this->rowHeight = $rowHeight;

        return $this;
    }

    public function getRowHeight()
    {
        return $this->rowHeight;
    }

    public function setTextWrap(bool $textWrap)
    {
        $this->textWrap = $textWrap;

        return $this;
    }

    public function getTextWrap(): bool
    {
        return $this->textWrap;
    }

    public function setCount(int $count)
    {
        $this->count = $count;
    }

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
