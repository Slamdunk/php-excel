<?php

declare(strict_types=1);

namespace Slam\Excel\Tests\Helper;

use ArrayIterator;
use org\bovigo\vfs;
use PHPExcel;
use PHPExcel_IOFactory;
use PHPUnit\Framework\TestCase;
use Slam\Excel;

final class TableWorkbookTest extends TestCase
{
    private $vfs;
    private $filename;

    protected function setUp()
    {
        $this->vfs = vfs\vfsStream::setup('root', 0770);
        $this->filename = vfs\vfsStream::url('root/test-encoding.xls');
    }

    public function testPostGenerationDetails()
    {
        $phpExcel = new Excel\Helper\TableWorkbook($this->filename);
        $activeSheet = $phpExcel->addWorksheet(\uniqid('Sheet_'));
        $table = new Excel\Helper\Table($activeSheet, 3, 4, \uniqid('Heading_'), new ArrayIterator([
            ['description' => 'AAA'],
            ['description' => 'BBB'],
        ]));

        $phpExcel->writeTable($table);
        $phpExcel->close();

        $this->assertSame(3, $table->getRowStart());
        $this->assertSame(7, $table->getRowEnd());

        $this->assertSame(5, $table->getDataRowStart());

        $this->assertSame(4, $table->getColumnStart());
        $this->assertSame(5, $table->getColumnEnd());

        $this->assertCount(2, $table);
        $this->assertSame(['description' => 'Description'], $table->getWrittenColumnTitles());
    }

    public function testHandleEncoding()
    {
        $textWithSpecialCharacters = \implode(' # ', [
            '€',
            'VIA MARTIRI DELLA LIBERTà 2',
            'FISSO20+OPZ.I¢CASA EURIB 3',
            'FISSO 20+ OPZIONE I°CASA EUR 6',
            '1° MAGGIO',
            'GIÀ XXXXXXX YYYYYYYYYYY',
            'FINANZIAMENTO 13/14¬ MENSILITà',

            'A \'\\|!"£$%&/()=?^àèìòùáéíóúÀÈÌÒÙÁÉÍÓÚ<>*ç°§[]@#{},.-;:_~` Z',
        ]);
        $heading = \sprintf('%s: %s', \uniqid('Heading_'), $textWithSpecialCharacters);
        $data = \sprintf('%s: %s', \uniqid('Data_'), $textWithSpecialCharacters);

        $phpExcel = new Excel\Helper\TableWorkbook($this->filename);
        $activeSheet = $phpExcel->addWorksheet(\uniqid());
        $table = new Excel\Helper\Table($activeSheet, 0, 0, $heading, new ArrayIterator([
            [
                'description' => $data,
            ],
        ]));

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

    public function testCellStyles()
    {
        $phpExcel = new Excel\Helper\TableWorkbook($this->filename);

        $columnCollection = new Excel\Helper\ColumnCollection([
            new Excel\Helper\Column('my_text', 'Foo1', 11, new Excel\Helper\CellStyle\Text()),
            new Excel\Helper\Column('my_perc', 'Foo2', 12, new Excel\Helper\CellStyle\Percentage()),
            new Excel\Helper\Column('my_inte', 'Foo3', 13, new Excel\Helper\CellStyle\Integer()),
            new Excel\Helper\Column('my_date', 'Foo4', 14, new Excel\Helper\CellStyle\Date()),
            new Excel\Helper\Column('my_amnt', 'Foo5', 15, new Excel\Helper\CellStyle\Amount()),
            new Excel\Helper\Column('my_itfc', 'Foo6', 16, new Excel\Helper\CellStyle\ItalianFiscalCode()),
            new Excel\Helper\Column('my_nodd', 'Foo7', 14, new Excel\Helper\CellStyle\Date()),
        ]);

        $activeSheet = $phpExcel->addWorksheet('names');
        $table = new Excel\Helper\Table($activeSheet, 1, 0, \uniqid('Heading_'), new ArrayIterator([
            [
                'my_text' => 'text',
                'my_perc' => 3.45,
                'my_inte' => 1234567.8,
                'my_date' => '2017-03-02',
                'my_amnt' => 1234567.89,
                'my_itfc' => 'AABB',
                'my_nodd' => null,
            ],
        ]));
        $table->setColumnCollection($columnCollection);

        $phpExcel->writeTable($table);
        $phpExcel->close();

        unset($phpExcel);

        $phpExcel = $this->getPhpExcelFromFile($this->filename);

        $firstSheet = $phpExcel->getSheet(0);
        $expected = [
            'A1' => null,
            'A2' => $table->getHeading(),

            'A3' => 'Foo1',
            'B3' => 'Foo2',
            'C3' => 'Foo3',
            'D3' => 'Foo4',
            'E3' => 'Foo5',
            'F3' => 'Foo6',

            'A4' => 'text',
            'B4' => 3.45,
            'C4' => 1234567.8,
            'D4' => '02/03/2017',
            'E4' => 1234567.89,
            'F4' => 'AABB',
        ];

        foreach ($expected as $cell => $content) {
            $this->assertSame($content, $firstSheet->getCell($cell)->getValue(), $cell);
        }
    }

    public function testTablePagination()
    {
        $phpExcel = new Excel\Helper\TableWorkbook($this->filename);
        $phpExcel->setRowsPerSheet(6);

        $activeSheet = $phpExcel->addWorksheet('names');
        $table = new Excel\Helper\Table($activeSheet, 1, 2, \uniqid(), new ArrayIterator([
            ['description' => 'AAA'],
            ['description' => 'BBB'],
            ['description' => 'CCC'],
            ['description' => 'DDD'],
            ['description' => 'EEE'],
        ]));

        $returnTable = $phpExcel->writeTable($table);
        $phpExcel->close();

        unset($phpExcel);

        $phpExcel = $this->getPhpExcelFromFile($this->filename);

        $firstSheet = $phpExcel->getSheet(0);
        $expected = [
            'C1' => null,
            'C2' => $table->getHeading(),
            'C3' => 'Description',
            'C4' => 'AAA',
            'C5' => 'BBB',
            'C6' => 'CCC',
            'C7' => null,
        ];

        foreach ($expected as $cell => $content) {
            $this->assertSame($content, $firstSheet->getCell($cell)->getValue());
        }

        $secondSheet = $phpExcel->getSheet(1);
        $expected = [
            'C1' => $returnTable->getHeading(),
            'C2' => 'Description',
            'C3' => 'DDD',
            'C4' => 'EEE',
            'C5' => null,
        ];

        foreach ($expected as $cell => $content) {
            $this->assertSame($content, $secondSheet->getCell($cell)->getValue());
        }

        $this->assertContains('names (', $firstSheet->getTitle());
        $this->assertContains('names (', $secondSheet->getTitle());
    }

    public function testEmptyTable()
    {
        $emptyTableMessage = \uniqid('no_data_');

        $phpExcel = new Excel\Helper\TableWorkbook($this->filename);
        $phpExcel->setEmptyTableMessage($emptyTableMessage);

        $activeSheet = $phpExcel->addWorksheet(\uniqid());
        $table = new Excel\Helper\Table($activeSheet, 0, 0, \uniqid(), new ArrayIterator([]));

        $phpExcel->writeTable($table);
        $phpExcel->close();

        unset($phpExcel);

        $phpExcel = $this->getPhpExcelFromFile($this->filename);

        $firstSheet = $phpExcel->getSheet(0);
        $expected = [
            'A1' => $table->getHeading(),
            'A2' => null,
            'A3' => $emptyTableMessage,
            'A4' => null,
        ];

        foreach ($expected as $cell => $content) {
            $this->assertSame($content, $firstSheet->getCell($cell)->getValue());
        }
    }

    public function testFontRowAttributesUsage()
    {
        $phpExcel = new Excel\Helper\TableWorkbook($this->filename);
        $activeSheet = $phpExcel->addWorksheet(\uniqid());
        $table = new Excel\Helper\Table($activeSheet, 0, 0, \uniqid(), new ArrayIterator([
            [
                'name' => 'Foo',
                'surname' => 'Bar',
            ],
            [
                'name' => 'Baz',
                'surname' => 'Xxx',
            ],
        ]));

        $table->setFontSize(12);
        $table->setRowHeight(33);
        $table->setTextWrap(true);

        $phpExcel->writeTable($table);
        $phpExcel->close();

        unset($phpExcel);

        $phpExcel = $this->getPhpExcelFromFile($this->filename);

        $firstSheet = $phpExcel->getSheet(0);
        $cell = $firstSheet->getCell('A3');
        $style = $cell->getStyle();

        $this->assertSame('Foo', $cell->getValue());
        $this->assertSame(12, $style->getFont()->getSize());
        $this->assertSame(33, $firstSheet->getRowDimension($cell->getRow())->getRowHeight());
        $this->assertTrue($style->getAlignment()->getWrapText());
    }

    /**
     * @dataProvider provideColumnStringFromIndexCases
     */
    public function testColumnStringFromIndex(int $index, string $columnString)
    {
        $this->assertSame($columnString, Excel\Helper\TableWorkbook::getColumnStringFromIndex($index));
    }

    public function provideColumnStringFromIndexCases()
    {
        return [
            [2, 'C'],
            [3, 'D'],
            [25, 'Z'],
            [26, 'AA'],
            [33, 'AH'],
            [701, 'ZZ'],
            [703, 'AAB'],
        ];
    }

    public function testColumnStringFromIndexExpectsPositiveValues()
    {
        $this->expectException(Excel\Exception\InvalidArgumentException::class);

        Excel\Helper\TableWorkbook::getColumnStringFromIndex(-1);
    }

    private function getPhpExcelFromFile(string $filename): PHPExcel
    {
        return PHPExcel_IOFactory::load($filename);
    }
}
