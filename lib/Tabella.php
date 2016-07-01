<?php

namespace Excel;

final class Tabella
{
    private $activeSheet;

    private $rigaIniziale;
    private $rigaMassima;
    private $rigaCorrente;

    private $colonnaIniziale;
    private $colonnaMassima;
    private $colonnaCorrente;

    private $intestazione;

    private $dati;

    private $colonnaCollection;

    private $bloccaRiquadri = true;

    private $isEmpty;

    public function __construct(Writer\Worksheet $activeSheet, $riga, $colonna, $intestazione, Iterator $dati)
    {
        $this->activeSheet = $activeSheet;

        $this->rigaIniziale =
        $this->rigaMassima =
        $this->rigaCorrente =
            (int) $riga
        ;

        $this->colonnaIniziale =
        $this->colonnaMassima =
        $this->colonnaCorrente =
            (int) $colonna
        ;

        $this->intestazione = (string) $intestazione;

        $this->dati = $dati;
    }

    public function getActiveSheet()
    {
        return $this->activeSheet;
    }

    public function getRigaIniziale()
    {
        return $this->rigaIniziale;
    }

    public function getRigaMassima()
    {
        return $this->rigaMassima;
    }

    public function getRigaCorrente()
    {
        return $this->rigaCorrente;
    }

    public function incrementaRiga()
    {
        ++$this->rigaCorrente;

        $this->rigaMassima = max($this->rigaMassima, $this->rigaCorrente);
    }

    public function getColonnaIniziale()
    {
        return $this->colonnaIniziale;
    }

    public function getColonnaMassima()
    {
        return $this->colonnaMassima;
    }

    public function getColonnaCorrente()
    {
        return $this->colonnaCorrente;
    }

    public function incrementaColonna()
    {
        ++$this->colonnaCorrente;

        $this->colonnaMassima = max($this->colonnaMassima, $this->colonnaCorrente);
    }

    public function ripristinaColonna()
    {
        $this->colonnaCorrente = $this->colonnaIniziale;
    }

    public function getIntestazione()
    {
        return $this->intestazione;
    }

    public function getDati()
    {
        return $this->dati;
    }

    public function setColonnaCollection(ColonnaCollection $colonnaCollection = null)
    {
        $this->colonnaCollection = $colonnaCollection;

        return $this;
    }

    public function getColonnaCollection()
    {
        return $this->colonnaCollection;
    }

    public function setBloccaRiquadri($bloccaRiquadri)
    {
        $this->bloccaRiquadri = (bool) $bloccaRiquadri;

        return $this;
    }

    public function getBloccaRiquadri()
    {
        return $this->bloccaRiquadri;
    }

    public function setEmpty($isEmpty)
    {
        $this->isEmpty = (bool) $isEmpty;
    }

    public function isEmpty()
    {
        if ($this->isEmpty === null) {
            throw new Excel\Exception\RuntimeException('Il workbook deve impostare il flag vuota sulla tabella');
        }

        return $this->isEmpty;
    }

    public function dividiTabellaSuNuovoSheet(Writer\Worksheet $activeSheet)
    {
        $nuovaTabella = new self($activeSheet, 0, $this->getColonnaIniziale(), $this->getIntestazione(), $this->getDati());
        $nuovaTabella->setColonnaCollection($this->getColonnaCollection());
        $nuovaTabella->setBloccaRiquadri($this->getBloccaRiquadri());

        return $nuovaTabella;
    }
}
