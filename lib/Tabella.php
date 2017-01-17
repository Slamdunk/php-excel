<?php

declare(strict_types=1);

namespace Excel;

use Countable;
use Iterator;

final class Tabella implements Countable
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

    private $count;

    public function __construct(Writer\Worksheet $activeSheet, int $riga, int $colonna, string $intestazione, Iterator $dati)
    {
        $this->activeSheet = $activeSheet;

        $this->rigaIniziale =
        $this->rigaMassima =
        $this->rigaCorrente =
            $riga
        ;

        $this->colonnaIniziale =
        $this->colonnaMassima =
        $this->colonnaCorrente =
            $colonna
        ;

        $this->intestazione = $intestazione;

        $this->dati = $dati;
    }

    public function getActiveSheet(): Writer\Worksheet
    {
        return $this->activeSheet;
    }

    public function getRigaIniziale(): int
    {
        return $this->rigaIniziale;
    }

    public function getRigaMassima(): int
    {
        return $this->rigaMassima;
    }

    public function getRigaCorrente(): int
    {
        return $this->rigaCorrente;
    }

    public function incrementaRiga()
    {
        ++$this->rigaCorrente;

        $this->rigaMassima = max($this->rigaMassima, $this->rigaCorrente);
    }

    public function getColonnaIniziale(): int
    {
        return $this->colonnaIniziale;
    }

    public function getColonnaMassima(): int
    {
        return $this->colonnaMassima;
    }

    public function getColonnaCorrente(): int
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

    public function getIntestazione(): string
    {
        return $this->intestazione;
    }

    public function getDati(): Iterator
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

    public function setBloccaRiquadri(bool $bloccaRiquadri)
    {
        $this->bloccaRiquadri = $bloccaRiquadri;

        return $this;
    }

    public function getBloccaRiquadri(): bool
    {
        return $this->bloccaRiquadri;
    }

    public function setCount(int $count)
    {
        $this->count = $count;
    }

    public function count()
    {
        if ($this->count === null) {
            throw new Exception\RuntimeException('Il workbook deve impostare il count sulla tabella');
        }

        return $this->count;
    }

    public function isEmpty(): bool
    {
        return $this->count() === 0;
    }

    public function dividiTabellaSuNuovoSheet(Writer\Worksheet $activeSheet): self
    {
        $nuovaTabella = new self($activeSheet, 0, $this->getColonnaIniziale(), $this->getIntestazione(), $this->getDati());
        $nuovaTabella->setColonnaCollection($this->getColonnaCollection());
        $nuovaTabella->setBloccaRiquadri($this->getBloccaRiquadri());

        return $nuovaTabella;
    }
}
