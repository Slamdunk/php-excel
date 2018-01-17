<?php

declare(strict_types=1);

namespace Slam\Excel\Helper\CellStyle;

use Slam\Excel\Helper\CellStyleInterface;
use Slam\Excel\Pear\Writer\Format;

final class Date implements CellStyleInterface
{
    public function decorateValue($value)
    {
        if (empty($value)) {
            return $value;
        }

        return \implode('/', \array_reverse(\explode('-', $value)));
    }

    public function styleCell(Format $format): void
    {
        $format->setAlign('center');
    }
}
