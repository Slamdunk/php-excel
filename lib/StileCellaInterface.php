<?php

namespace Excel;

interface StileCellaInterface
{
    public function decorateValue($value);

    public function styleCell(Writer\Format $format);
}
