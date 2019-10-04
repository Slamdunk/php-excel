<?php

declare(strict_types=1);

namespace Slam\Excel\Tests\Helper;

use ArrayIterator;
use PHPUnit\Framework\TestCase;
use Slam\Excel;

final class TableTest extends TestCase
{
    /**
     * @var Excel\Pear\Writer\Workbook
     */
    private $phpExcel;

    /**
     * @var Excel\Pear\Writer\Worksheet
     */
    private $activeSheet;

    /**
     * @var ArrayIterator
     */
    private $data;

    /**
     * @var Excel\Helper\Table
     */
    private $table;

    protected function setUp()
    {
        $this->phpExcel    = new Excel\Pear\Writer\Workbook(\uniqid());
        $this->activeSheet = $this->phpExcel->addWorksheet('sheet');

        $this->data = new ArrayIterator(['a', 'b']);

        $this->table = new Excel\Helper\Table($this->activeSheet, 3, 12, 'My Heading', $this->data);
    }

    public function testRowAndColumn()
    {
        static::assertSame($this->activeSheet, $this->table->getActiveSheet());
        static::assertSame('My Heading', $this->table->getHeading());
        static::assertSame($this->data, $this->table->getData());

        static::assertNull($this->table->getDataRowStart());

        $this->table->incrementRow();
        $this->table->flagDataRowStart();
        $this->table->incrementRow();

        static::assertSame(3, $this->table->getRowStart());
        static::assertSame(4, $this->table->getDataRowStart());
        static::assertSame(5, $this->table->getRowEnd());
        static::assertSame(5, $this->table->getRowCurrent());

        $this->table->incrementColumn();
        $this->table->incrementColumn();

        static::assertSame(12, $this->table->getColumnStart());
        static::assertSame(14, $this->table->getColumnEnd());
        static::assertSame(14, $this->table->getColumnCurrent());

        $this->table->resetColumn();

        static::assertSame(12, $this->table->getColumnStart());
        static::assertSame(14, $this->table->getColumnEnd());
        static::assertSame(12, $this->table->getColumnCurrent());

        $this->table->setCount(0);
        static::assertCount(0, $this->table);
        static::assertTrue($this->table->isEmpty());

        $this->table->setCount(5);
        static::assertCount(5, $this->table);
        static::assertFalse($this->table->isEmpty());

        static::assertTrue($this->table->getFreezePanes());
        $this->table->setFreezePanes(false);
        static::assertFalse($this->table->getFreezePanes());

        static::assertNull($this->table->getWrittenColumnTitles());
        $columns = [
            'column_1' => 'Name',
            'column_2' => 'Surname',
        ];
        $this->table->setWrittenColumnTitles($columns);
        static::assertSame($columns, $this->table->getWrittenColumnTitles());
    }

    public function testTableCountMustBeSet()
    {
        $this->expectException(Excel\Exception\RuntimeException::class);

        $this->table->count();
    }

    public function testSplitTableIfNeeded()
    {
        $newSheet = $this->phpExcel->addWorksheet('sheet 2');
        $this->table->setFreezePanes(false);
        $newTable = $this->table->splitTableOnNewWorksheet($newSheet);

        static::assertNotSame($this->table, $newTable);

        // The starting row must be the first of the new sheet
        static::assertSame(0, $newTable->getRowStart());
        static::assertSame(0, $newTable->getRowEnd());
        static::assertSame(0, $newTable->getRowCurrent());

        // The starting column must be the same of the previous sheet
        static::assertSame(12, $newTable->getColumnStart());
        static::assertSame(12, $newTable->getColumnEnd());
        static::assertSame(12, $newTable->getColumnCurrent());

        static::assertSame($this->table->getFreezePanes(), $newTable->getFreezePanes());
    }

    public function testFontRowAttributes()
    {
        static::assertSame(8, $this->table->getFontSize());
        static::assertNull($this->table->getRowHeight());
        static::assertFalse($this->table->getTextWrap());

        $this->table->setFontSize($fontSize = \mt_rand(10, 100));
        $this->table->setRowHeight($rowHeight = \mt_rand(10, 100));
        $this->table->setTextWrap(true);

        static::assertSame($fontSize, $this->table->getFontSize());
        static::assertSame($rowHeight, $this->table->getRowHeight());
        static::assertTrue($this->table->getTextWrap());
    }
}
