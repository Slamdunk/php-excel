<?php

declare(strict_types=1);

namespace Slam\Excel\Helper;

final class Column implements ColumnInterface
{
    private $key;

    private $heading;

    private $width;

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
