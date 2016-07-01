<?php

namespace Excel\StileCella;

use Excel;

final class CodiceFiscale implements Excel\StileCellaInterface
{
    public function decorateValue($value)
    {
        return $value;
    }

    public function styleCell(Excel\Writer\Format $format)
    {
        $format->setNumFormat('00000000000');
        $format->setAlign('left');
    }
}
