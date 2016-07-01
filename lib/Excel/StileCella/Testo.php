<?php

final class Excel_StileCella_Testo implements Excel_StileCellaInterface
{
    public function decorateValue($value)
    {
        return $value;
    }

    public function styleCell(Excel_Writer_Format $format)
    {
        $format->setAlign('left');
    }
}
