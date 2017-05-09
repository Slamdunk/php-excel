<?php

declare(strict_types=1);

namespace Slam\Excel\Helper\CellStyle;

use Slam\Excel\Helper\CellStyleInterface;
use Slam\Excel\Pear\Writer\Format;

final class Text implements CellStyleInterface
{
    public function decorateValue($value)
    {
        return $value;
    }

    public function styleCell(Format $format)
    {
        $format->setAlign('left');
    }
}
