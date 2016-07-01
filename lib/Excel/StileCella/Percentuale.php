<?php

final class Excel_StileCella_Percentuale implements Excel_StileCellaInterface
{
    public function decorateValue($value)
    {
        return $value;
    }

    public function styleCell(Excel_Writer_Format $format)
    {
        $format->setNumFormat('#,##0.000');
    }
}
