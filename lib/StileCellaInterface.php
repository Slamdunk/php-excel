<?php

declare(strict_types=1);

namespace Excel;

interface StileCellaInterface
{
    public function decorateValue($value);

    public function styleCell(Writer\Format $format);
}
