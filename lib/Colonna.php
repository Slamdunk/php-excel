<?php

namespace Excel;

final class Colonna implements ColonnaInterface
{
    private $chiave;

    private $intestazione;

    private $largezza;

    private $stileCella;

    public function __construct($chiave, $intestazione, $largezza, StileCellaInterface $stileCella)
    {
        $this->chiave       = $chiave;
        $this->intestazione = $intestazione;
        $this->largezza     = $largezza;
        $this->stileCella   = $stileCella;
    }

    public function getChiave()
    {
        return $this->chiave;
    }

    public function getIntestazione()
    {
        return $this->intestazione;
    }

    public function getLarghezza()
    {
        return $this->largezza;
    }

    public function getStileCella()
    {
        return $this->stileCella;
    }
}
