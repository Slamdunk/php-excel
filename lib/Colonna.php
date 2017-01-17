<?php

declare(strict_types=1);

namespace Excel;

final class Colonna implements ColonnaInterface
{
    private $chiave;

    private $intestazione;

    private $largezza;

    private $stileCella;

    public function __construct(string $chiave, string $intestazione, int $largezza, StileCellaInterface $stileCella)
    {
        $this->chiave       = $chiave;
        $this->intestazione = $intestazione;
        $this->largezza     = $largezza;
        $this->stileCella   = $stileCella;
    }

    public function getChiave(): string
    {
        return $this->chiave;
    }

    public function getIntestazione(): string
    {
        return $this->intestazione;
    }

    public function getLarghezza(): int
    {
        return $this->largezza;
    }

    public function getStileCella(): StileCellaInterface
    {
        return $this->stileCella;
    }
}
