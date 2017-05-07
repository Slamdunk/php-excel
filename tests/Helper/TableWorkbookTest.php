<?php

declare(strict_types=1);

namespace Slam\Excel\Tests\Helper;

use ArrayIterator;
use Slam\Excel;
use org\bovigo\vfs;
use PHPExcel_IOFactory;
use PHPUnit\Framework\TestCase;

final class TableWorkbookTest extends TestCase
{
    protected function setUp()
    {
        $this->vfs = vfs\vfsStream::setup('root', 0770);
        $this->filename = vfs\vfsStream::url('root/test-encoding.xls');
    }

    public function testHandleEncoding()
    {
        $textWithSpecialCharacters = implode(' # ', array(
            '€',
            'VIA MARTIRI DELLA LIBERTà 2',
            'FISSO20+OPZ.I¢CASA EURIB 3',
            'FISSO 20+ OPZIONE I°CASA EUR 6',
            '1° MAGGIO',
            'GIÀ XXXXXXX YYYYYYYYYYY',
            'FINANZIAMENTO 13/14¬ MENSILITà',

            'A \'\\|!"£$%&/()=?^àèìòùáéíóúÀÈÌÒÙÁÉÍÓÚ<>*ç°§[]@#{},.-;:_~` Z',
        ));
        $heading = sprintf('%s: %s', uniqid('Heading_'), $textWithSpecialCharacters);
        $data = sprintf('%s: %s', uniqid('Data_'), $textWithSpecialCharacters);

        $phpExcel = new Excel\Helper\TableWorkbook($this->filename);
        $activeSheet = $phpExcel->addWorksheet(uniqid());
        $table = new Excel\Helper\Table($activeSheet, 0, 0, $heading, new ArrayIterator(array(
            array(
                'description' => $data,
            ),
        )));

        $phpExcel->writeTable($table);
        $phpExcel->close();

        unset($phpExcel);

        $phpExcel = $this->getPhpExcelFromFile($this->filename);
        $firstSheet = $phpExcel->getActiveSheet();
        $this->assertSame($activeSheet->getName(), $firstSheet->getTitle());

        // Heading
        $value = $firstSheet->getCell('A1')->getValue();
        $this->assertSame($heading, $value);

        // Data
        $value = $firstSheet->getCell('A3')->getValue();
        $this->assertSame($data, $value);
    }

    public function testTablePagination()
    {
        $phpExcel = new Excel\Helper\TableWorkbook($this->filename);
        $phpExcel->setRowsPerSheet(6);

        $activeSheet = $phpExcel->addWorksheet('names');
        $table = new Excel\Helper\Table($activeSheet, 1, 2, uniqid(), new ArrayIterator(array(
            array('description' => 'AAA'),
            array('description' => 'BBB'),
            array('description' => 'CCC'),
            array('description' => 'DDD'),
            array('description' => 'EEE'),
        )));

        $returnTable = $phpExcel->writeTable($table);
        $phpExcel->close();

        unset($phpExcel);

        $phpExcel = $this->getPhpExcelFromFile($this->filename);

        $firstSheet = $phpExcel->getSheet(0);
        $expected = array(
            'C1' => null,
            'C2' => $table->getHeading(),
            'C3' => 'Description',
            'C4' => 'AAA',
            'C5' => 'BBB',
            'C6' => 'CCC',
            'C7' => null,
        );

        foreach ($expected as $cell => $content) {
            $this->assertSame($content, $firstSheet->getCell($cell)->getValue());
        }

        $secondSheet = $phpExcel->getSheet(1);
        $expected = array(
            'C1' => $returnTable->getHeading(),
            'C2' => 'Description',
            'C3' => 'DDD',
            'C4' => 'EEE',
            'C5' => null,
        );

        foreach ($expected as $cell => $content) {
            $this->assertSame($content, $secondSheet->getCell($cell)->getValue());
        }

        $this->assertContains('names (', $firstSheet->getTitle());
        $this->assertContains('names (', $secondSheet->getTitle());
    }

    public function testEmptyTable()
    {
        $emptyTableMessage = uniqid('no_data_');
        
        $phpExcel = new Excel\Helper\TableWorkbook($this->filename);
        $phpExcel->setEmptyTableMessage($emptyTableMessage);

        $activeSheet = $phpExcel->addWorksheet(uniqid());
        $table = new Excel\Helper\Table($activeSheet, 0, 0, uniqid(), new ArrayIterator(array()));

        $phpExcel->writeTable($table);
        $phpExcel->close();

        unset($phpExcel);

        $phpExcel = $this->getPhpExcelFromFile($this->filename);

        $firstSheet = $phpExcel->getSheet(0);
        $expected = array(
            'A1' => $table->getHeading(),
            'A2' => null,
            'A3' => $emptyTableMessage,
            'A4' => null,
        );

        foreach ($expected as $cell => $content) {
            $this->assertSame($content, $firstSheet->getCell($cell)->getValue());
        }
    }

    private function getPhpExcelFromFile(string $filename)
    {
        return PHPExcel_IOFactory::load($filename);
    }
}
