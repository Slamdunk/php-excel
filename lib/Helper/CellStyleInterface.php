<?php

declare(strict_types=1);

namespace Slam\Excel\Helper;

use Slam\Excel\Pear\Writer\Format;

interface CellStyleInterface
{
    /**
     * @param mixed $value
     *
     * @return mixed
     */
    public function decorateValue($value);

    public function styleCell(Format $format): void;
}
