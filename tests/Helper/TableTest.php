<?php

declare(strict_types=1);

namespace Slam\Excel\Tests\Helper;

use ArrayIterator;
use PHPUnit\Framework\TestCase;
use Slam\Excel;

final class TableTest extends TestCase
{
    protected function setUp()
    {
        $this->phpExcel = new Excel\Pear\Writer\Workbook(uniqid());
        $this->activeSheet = $this->phpExcel->addWorksheet('sheet');

        $this->data = new ArrayIterator(array('a', 'b'));

        $this->table = new Excel\Helper\Table($this->activeSheet, 3, 12, 'My Heading', $this->data);
    }

    public function testRowAndColumn()
    {
        $this->assertSame($this->activeSheet, $this->table->getActiveSheet());
        $this->assertSame('My Heading', $this->table->getHeading());
        $this->assertSame($this->data, $this->table->getData());

        $this->table->incrementRow();
        $this->table->incrementRow();

        $this->assertSame(3, $this->table->getRowStart());
        $this->assertSame(5, $this->table->getRowEnd());
        $this->assertSame(5, $this->table->getRowCurrent());

        $this->table->incrementColumn();
        $this->table->incrementColumn();

        $this->assertSame(12, $this->table->getColumnStart());
        $this->assertSame(14, $this->table->getColumnEnd());
        $this->assertSame(14, $this->table->getColumnCurrent());

        $this->table->resetColumn();

        $this->assertSame(12, $this->table->getColumnStart());
        $this->assertSame(14, $this->table->getColumnEnd());
        $this->assertSame(12, $this->table->getColumnCurrent());

        $this->table->setCount(0);
        $this->assertCount(0, $this->table);
        $this->assertTrue($this->table->isEmpty());

        $this->table->setCount(5);
        $this->assertCount(5, $this->table);
        $this->assertFalse($this->table->isEmpty());

        $this->assertTrue($this->table->getFreezePanes());
        $this->table->setFreezePanes(false);
        $this->assertFalse($this->table->getFreezePanes());
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

        $this->assertNotSame($this->table, $newTable);

        // The starting row must be the first of the new sheet
        $this->assertSame(0, $newTable->getRowStart());
        $this->assertSame(0, $newTable->getRowEnd());
        $this->assertSame(0, $newTable->getRowCurrent());

        // The starting column must be the same of the previous sheet
        $this->assertSame(12, $newTable->getColumnStart());
        $this->assertSame(12, $newTable->getColumnEnd());
        $this->assertSame(12, $newTable->getColumnCurrent());

        $this->assertSame($this->table->getFreezePanes(), $newTable->getFreezePanes());
    }

    public function testFontRowAttributes()
    {
        $this->assertSame(8, $this->table->getFontSize());
        $this->assertSame(null, $this->table->getRowHeight());
        $this->assertSame(false, $this->table->getTextWrap());

        $this->table->setFontSize($fontSize = mt_rand(10, 100));
        $this->table->setRowHeight($rowHeight = mt_rand(10, 100));
        $this->table->setTextWrap(true);

        $this->assertSame($fontSize, $this->table->getFontSize());
        $this->assertSame($rowHeight, $this->table->getRowHeight());
        $this->assertSame(true, $this->table->getTextWrap());
    }
}
