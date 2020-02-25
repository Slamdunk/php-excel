<?php

declare(strict_types=1);

namespace Slam\Excel\Helper;

use Slam\Excel\Exception;

final class ColumnCollection implements ColumnCollectionInterface
{
    /**
     * @var array<string, ColumnInterface>
     */
    private $columns = [];

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
    public function offsetSet($offset, $value)
    {
        throw new Exception\RuntimeException('Collection not editable');
    }

    /**
     * @param string $offset
     *
     * @return bool
     */
    public function offsetExists($offset)
    {
        return isset($this->columns[$offset]);
    }

    /**
     * @param string $offset
     */
    public function offsetUnset($offset)
    {
        throw new Exception\RuntimeException('Collection not editable');
    }

    /**
     * @param string $offset
     *
     * @return null|mixed
     */
    public function offsetGet($offset)
    {
        return $this->columns[$offset] ?? null;
    }
}
