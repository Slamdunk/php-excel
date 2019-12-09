<?php

declare(strict_types=1);

namespace Slam\Excel\Tests\Helper;

use PHPUnit\Framework\TestCase;
use Slam\Excel\Exception;
use Slam\Excel\Helper;

final class ColumnCollectionTest extends TestCase
{
    private $column;
    private $collection;

    protected function setUp(): void
    {
        $this->column = new Helper\Column('foo', 'Foo', 10, new Helper\CellStyle\Text());

        $this->collection = new Helper\ColumnCollection([
            $this->column,
        ]);
    }

    public function testBaseFunctionalities(): void
    {
        self::assertArrayHasKey('foo', $this->collection);
        self::assertSame($this->column, $this->collection['foo']);
    }

    public function testNotEditableOnSet(): void
    {
        $this->expectException(Exception\RuntimeException::class);

        $this->collection['foo'] = 1;
    }

    public function testNotEditableOnUnset(): void
    {
        $this->expectException(Exception\RuntimeException::class);

        unset($this->collection['foo']);
    }
}
