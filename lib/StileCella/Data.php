<?php

namespace Excel\StileCella;

use Excel;

final class Data implements Excel\StileCellaInterface
{
    public function decorateValue($value)
    {
        if (empty($value)) {
            return $value;
        }

        return implode('/', array_reverse(explode('-', $value)));
    }

    public function styleCell(Excel\Writer\Format $format)
    {
        $format->setAlign('center');
    }
}
