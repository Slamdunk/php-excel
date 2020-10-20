<?php

declare(strict_types=1);

namespace Slam\Excel\Helper;

use Slam\Excel\Exception;

final class ColumnCollection implements ColumnCollectionInterface
{
    /**
     * @var array<string, ColumnInterface>
     */
    private array $columns = [];

    public function __construct(array $columns)
    {
        foreach ($columns as $column) {
            $this->addColumn($column);
        }
    }

    private function addColumn(ColumnInterface $column): void
    {
        $this->columns[$column->getKey()] = $column;
    }

    /**
     * @param string $offset
     * @param mixed  $value
     */
    public function offsetSet($offset, $value): void
    {
        throw new Exception\RuntimeException('Collection not editable');
    }

    /**
     * @param string $offset
     */
    public function offsetExists($offset): bool
    {
        return isset($this->columns[$offset]);
    }

    /**
     * @param string $offset
     */
    public function offsetUnset($offset): void
    {
        throw new Exception\RuntimeException('Collection not editable');
    }

    /**
     * @param string $offset
     */
    public function offsetGet($offset): ?ColumnInterface
    {
        return $this->columns[$offset] ?? null;
    }
}
