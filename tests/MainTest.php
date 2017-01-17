<?php

declare(strict_types=1);

namespace ExcelTest;

use Excel;
use org\bovigo\vfs;
use PHPUnit_Framework_TestCase;

final class MainTest extends PHPUnit_Framework_TestCase
{
    protected function setUp()
    {
        Excel\OLE::$gmmktime = gmmktime(1, 1, 1, 1, 1, 2000);

        $this->vfs = vfs\vfsStream::setup('root', 0770);
        $this->filename = vfs\vfsStream::url('root/test.xls');

        // $this->filename = __DIR__ . '/stock.xls';

        $this->xls = new Excel\Writer\Workbook($this->filename);
    }

    public function testGenerazioneFileBase()
    {
        $this->xls->setCustomColor(60, hexdec('7f'), hexdec('7f'), hexdec('7f'));
        $this->xls->setCustomColor(61, hexdec('e8'), hexdec('e8'), hexdec('e8'));
        $this->xls->setCustomColor(62, hexdec('cc'), hexdec('cc'), hexdec('cc'));

        $sheet = $this->xls->addWorksheet('FoglioCustom');
        $sheet->setLandscape();
        $sheet->setMargins(0.2);
        $sheet->hideGridLines();
        $sheet->setColumn(1, 1, 20);

        $header = $this->xls->addFormat();
        $header->setColor('yellow');
        $header->setBold();
        $header->setSize(15);
        $header->setAlign('center');
        $header->setBgColor(61);

        $sheet->write(1, 1, 'LCR', $header);

        $sheet->freezePanes(array(2, 0));

        $sheet->writeString(2, 1, '0123');

        $format2 = $this->xls->addFormat();
        $format2->setColor('red');
        $format2->setItalic();
        $format2->setBorder(2);
        $format2->setBorderColor('lime');
        $format2->setNumFormat('#,##0.00');
        $format2->setFgColor(62);
        $format2->setAlign('top');

        $sheet->write(3, 2, '95000', $format2);

        $sheet->write(7, 3, 1.1);
        $sheet->write(8, 3, 2);
        $sheet->writeFormula(9, 3, sprintf('=SUM(%s:%s)', $this->xls->rowcolToCell(7, 3), $this->xls->rowcolToCell(8, 3)));

        $this->xls->close();

        $this->assertLessThan(Excel\OLE::Excel_OLE_DATA_SIZE_SMALL, filesize($this->filename));
        $this->assertSame('43093f883818e44f4dd62f0382d95ecf6b689004', hash('sha1', file_get_contents($this->filename)));
    }

    public function testFileGrandi()
    {
        $sheet = $this->xls->addWorksheet('FoglioCustom');

        for ($i = 0; $i < 1000; ++$i) {
            $sheet->writeString($i, 1, 'foobar' . $i);
        }

        $this->xls->close();

        $this->assertGreaterThan(Excel\OLE::Excel_OLE_DATA_SIZE_SMALL, filesize($this->filename));
        $this->assertSame('884f515114705011286817afff99f848f76b0ce2', hash('sha1', file_get_contents($this->filename)));
    }

    /**
     * @dataProvider dataProviderTestIndiceColonnaInNumero
     */
    public function testIndiceColonnaInNumero($indice, $lettera)
    {
        $this->assertSame($lettera . '2', $this->xls->rowcolToCell(1, $indice));

        $this->setExpectedException('Excel\Exception\InvalidArgumentException');

        $this->xls->rowcolToCell(1, 50000);
    }

    public function dataProviderTestIndiceColonnaInNumero()
    {
        return array(
            array(0, 'A'),
            array(30, 'AE'),
            array(231, 'HX'),
        );
    }
}
