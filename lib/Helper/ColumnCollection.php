<?php

declare(strict_types=1);

namespace Slam\Excel\Helper;

use Slam\Excel\Exception;

final class ColumnCollection implements ColumnCollectionInterface
{
    private $columns = array();

    public function __construct(array $columns)
    {
        foreach ($columns as $column) {
            $this->addColumn($column);
        }
    }

    private function addColumn(ColumnInterface $column)
    {
        $this->columns[$column->getKey()] = $column;
    }

    public function offsetSet($offset, $value)
    {
        throw new Exception\RuntimeException('Collection not editable');
    }

    public function offsetExists($offset)
    {
        return isset($this->columns[$offset]);
    }

    public function offsetUnset($offset)
    {
        throw new Exception\RuntimeException('Collection not editable');
    }

    public function offsetGet($offset)
    {
        return $this->columns[$offset] ?? null;
    }
}
