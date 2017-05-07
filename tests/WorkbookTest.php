<?php

declare(strict_types=1);

namespace ExcelTest;

use ArrayIterator;
use Excel;
use org\bovigo\vfs;
use PHPExcel_IOFactory;
use PHPUnit\Framework\TestCase;

final class WorkbookTest extends TestCase
{
    protected function setUp()
    {
        $this->vfs = vfs\vfsStream::setup('root', 0770);
        $this->filename = vfs\vfsStream::url('root/test-encoding.xls');
    }

    public function testGestisceCorrettamenteEncoding()
    {
        $testoConCaratteriSpeciali = implode(' # ', array(
            '€',
            'VIA MARTIRI DELLA LIBERTà 2',
            'FISSO20+OPZ.I¢CASA EURIB 3',
            'FISSO 20+ OPZIONE I°CASA EUR 6',
            '1° MAGGIO',
            'GIÀ XXXXXXX YYYYYYYYYYY',
            'FINANZIAMENTO 13/14¬ MENSILITà',

            'A \'\\|!"£$%&/()=?^àèìòùáéíóúÀÈÌÒÙÁÉÍÓÚ<>*ç°§[]@#{},.-;:_~` Z',
        ));
        $intestazione = sprintf('%s: %s', uniqid('Intestazione_'), $testoConCaratteriSpeciali);
        $dati = sprintf('%s: %s', uniqid('Dati_'), $testoConCaratteriSpeciali);

        $phpExcel = new Excel\Workbook($this->filename);
        $activeSheet = $phpExcel->addWorksheet(uniqid());
        $tabella = new Excel\Tabella($activeSheet, 0, 0, $intestazione, new ArrayIterator(array(
            array(
                'descrizione' => $dati,
            ),
        )));

        $phpExcel->scriviTabella($tabella);
        $phpExcel->close();

        unset($phpExcel);

        $phpExcel = $this->getPhpExcelFromFile($this->filename);
        $primaScheda = $phpExcel->getActiveSheet();
        $this->assertSame($activeSheet->getName(), $primaScheda->getTitle());

        // Intestazione
        $value = $primaScheda->getCell('A1')->getValue();
        $this->assertSame($intestazione, $value);

        // Dati
        $value = $primaScheda->getCell('A3')->getValue();
        $this->assertSame($dati, $value);
    }

    public function testPaginazioneTabellaTraPiuSchede()
    {
        $phpExcel = new Excel\Workbook($this->filename);
        $phpExcel->setRighePerPagina(6);

        $activeSheet = $phpExcel->addWorksheet('finanziamenti');
        $tabella = new Excel\Tabella($activeSheet, 1, 2, uniqid(), new ArrayIterator(array(
            array('descrizione' => 'AAA'),
            array('descrizione' => 'BBB'),
            array('descrizione' => 'CCC'),
            array('descrizione' => 'DDD'),
            array('descrizione' => 'EEE'),
        )));

        $tabellaDiRitorno = $phpExcel->scriviTabella($tabella);
        $phpExcel->close();

        unset($phpExcel);

        $phpExcel = $this->getPhpExcelFromFile($this->filename);

        $primaScheda = $phpExcel->getSheet(0);
        $contenutiAttesi = array(
            'C1' => null,
            'C2' => $tabella->getIntestazione(),
            'C3' => 'Descrizione',
            'C4' => 'AAA',
            'C5' => 'BBB',
            'C6' => 'CCC',
            'C7' => null,
        );

        foreach ($contenutiAttesi as $cella => $contenuto) {
            $this->assertSame($contenuto, $primaScheda->getCell($cella)->getValue());
        }

        $secondaScheda = $phpExcel->getSheet(1);
        $contenutiAttesi = array(
            'C1' => $tabellaDiRitorno->getIntestazione(),
            'C2' => 'Descrizione',
            'C3' => 'DDD',
            'C4' => 'EEE',
            'C5' => null,
        );

        foreach ($contenutiAttesi as $cella => $contenuto) {
            $this->assertSame($contenuto, $secondaScheda->getCell($cella)->getValue());
        }

        $this->assertContains('finanziamenti (', $primaScheda->getTitle());
        $this->assertContains('finanziamenti (', $secondaScheda->getTitle());
    }

    public function testTabellaVuota()
    {
        $phpExcel = new Excel\Workbook($this->filename);

        $activeSheet = $phpExcel->addWorksheet(uniqid());
        $tabella = new Excel\Tabella($activeSheet, 0, 0, uniqid(), new ArrayIterator(array()));

        $phpExcel->scriviTabella($tabella);
        $phpExcel->close();

        unset($phpExcel);

        $phpExcel = $this->getPhpExcelFromFile($this->filename);

        $primaScheda = $phpExcel->getSheet(0);
        $contenutiAttesi = array(
            'A1' => $tabella->getIntestazione(),
            'A2' => null,
            'A3' => 'Nessun dato per questa estrazione',
            'A4' => null,
        );

        foreach ($contenutiAttesi as $cella => $contenuto) {
            $this->assertSame($contenuto, $primaScheda->getCell($cella)->getValue());
        }
    }

    private function getPhpExcelFromFile(string $filename)
    {
        return PHPExcel_IOFactory::load($filename);
    }
}
