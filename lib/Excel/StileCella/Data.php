<?php

final class Excel_StileCella_Data implements Excel_StileCellaInterface
{
    public function decorateValue($value)
    {
        if (empty($value)) {
            return $value;
        }

        return implode('/', array_reverse(explode('-', $value)));
    }

    public function styleCell(Excel_Writer_Format $format)
    {
        $format->setAlign('center');
    }
}
