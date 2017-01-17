<?php

declare(strict_types=1);

namespace Excel;

interface ColonnaInterface
{
    public function getChiave(): string;

    public function getIntestazione(): string;

    public function getLarghezza(): int;

    public function getStileCella(): StileCellaInterface;
}
