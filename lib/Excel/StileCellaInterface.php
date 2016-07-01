<?php

interface Excel_StileCellaInterface
{
    public function decorateValue($value);

    public function styleCell(Excel_Writer_Format $format);
}
