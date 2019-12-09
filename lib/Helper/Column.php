<?php

declare(strict_types=1);

namespace Slam\Excel\Helper;

final class Column implements ColumnInterface
{
    /**
     * @var string
     */
    private $key;

    /**
     * @var string
     */
    private $heading;

    /**
     * @var int
     */
    private $width;

    /**
     * @var CellStyleInterface
     */
    private $cellStyle;

    public function __construct(string $key, string $heading, int $width, CellStyleInterface $cellStyle)
    {
        $this->key          = $key;
        $this->heading      = $heading;
        $this->width        = $width;
        $this->cellStyle    = $cellStyle;
    }

    public function getKey(): string
    {
        return $this->key;
    }

    public function getHeading(): string
    {
        return $this->heading;
    }

    public function getWidth(): int
    {
        return $this->width;
    }

    public function getCellStyle(): CellStyleInterface
    {
        return $this->cellStyle;
    }
}
