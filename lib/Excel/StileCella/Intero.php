<?php

final class Excel_StileCella_Intero implements Excel_StileCellaInterface
{
    public function decorateValue($value)
    {
        return $value;
    }

    public function styleCell(Excel_Writer_Format $format)
    {
        $format->setNumFormat('#,##0');
        $format->setAlign('center');
    }
}
