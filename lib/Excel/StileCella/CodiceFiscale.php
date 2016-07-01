<?php

final class Excel_StileCella_CodiceFiscale implements Excel_StileCellaInterface
{
    public function decorateValue($value)
    {
        return $value;
    }

    public function styleCell(Excel_Writer_Format $format)
    {
        $format->setNumFormat('00000000000');
        $format->setAlign('left');
    }
}
