<?php

declare(strict_types=1);

namespace ExcelTest;

use ArrayIterator;
use Excel;
use PHPUnit\Framework\TestCase;

final class TabellaTest extends TestCase
{
    protected function setUp()
    {
        $this->phpExcel = new Excel\Writer\Workbook(uniqid());
        $this->activeSheet = $this->phpExcel->addWorksheet('scheda');

        $this->dati = new ArrayIterator(array('a', 'b'));

        $this->tabella = new Excel\Tabella($this->activeSheet, 3, 12, 'Intestazione', $this->dati);
    }

    public function testRigaColonna()
    {
        $this->assertSame($this->activeSheet, $this->tabella->getActiveSheet());
        $this->assertSame('Intestazione', $this->tabella->getIntestazione());
        $this->assertSame($this->dati, $this->tabella->getDati());

        $this->tabella->incrementaRiga();
        $this->tabella->incrementaRiga();

        $this->assertSame(3, $this->tabella->getRigaIniziale());
        $this->assertSame(5, $this->tabella->getRigaMassima());
        $this->assertSame(5, $this->tabella->getRigaCorrente());

        $this->tabella->incrementaColonna();
        $this->tabella->incrementaColonna();

        $this->assertSame(12, $this->tabella->getColonnaIniziale());
        $this->assertSame(14, $this->tabella->getColonnaMassima());
        $this->assertSame(14, $this->tabella->getColonnaCorrente());

        $this->tabella->ripristinaColonna();

        $this->assertSame(12, $this->tabella->getColonnaIniziale());
        $this->assertSame(14, $this->tabella->getColonnaMassima());
        $this->assertSame(12, $this->tabella->getColonnaCorrente());

        $this->tabella->setCount(0);
        $this->assertCount(0, $this->tabella);
        $this->assertTrue($this->tabella->isEmpty());

        $this->tabella->setCount(5);
        $this->assertCount(5, $this->tabella);
        $this->assertFalse($this->tabella->isEmpty());

        $this->assertTrue($this->tabella->getBloccaRiquadri());
        $this->tabella->setBloccaRiquadri(false);
        $this->assertFalse($this->tabella->getBloccaRiquadri());
    }

    public function testTabellaDeveAvereFlagVuota()
    {
        $this->expectException('Excel\Exception\RuntimeException');

        $this->tabella->count();
    }

    public function testSplitTabellaSeServePerPaginazione()
    {
        $newSheet = $this->phpExcel->addWorksheet('scheda 2');
        $this->tabella->setBloccaRiquadri(false);
        $tabellaNuova = $this->tabella->dividiTabellaSuNuovoSheet($newSheet);

        $this->assertNotSame($this->tabella, $tabellaNuova);

        // La riga iniziale deve essere la prima riga del nuovo foglio
        $this->assertSame(0, $tabellaNuova->getRigaIniziale());
        $this->assertSame(0, $tabellaNuova->getRigaMassima());
        $this->assertSame(0, $tabellaNuova->getRigaCorrente());

        // La colonna iniziale invece deve essere la stessa del foglio precedente
        $this->assertSame(12, $tabellaNuova->getColonnaIniziale());
        $this->assertSame(12, $tabellaNuova->getColonnaMassima());
        $this->assertSame(12, $tabellaNuova->getColonnaCorrente());

        $this->assertSame($this->tabella->getBloccaRiquadri(), $tabellaNuova->getBloccaRiquadri());
    }
}
