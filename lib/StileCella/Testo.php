<?php

namespace Excel\StileCella;

use Excel;

final class Testo implements Excel\StileCellaInterface
{
    public function decorateValue($value)
    {
        return $value;
    }

    public function styleCell(Excel\Writer\Format $format)
    {
        $format->setAlign('left');
    }
}
